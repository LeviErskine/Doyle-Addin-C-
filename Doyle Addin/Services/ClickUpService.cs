#nullable enable
namespace DoyleAddin.Services;

using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Options;

/// <summary>
///     Service for creating ClickUp tasks, subtasks, and checklist items during DXF exports.
///     Implements rate limiting for Business Plus plan (1000 requests/minute).
/// </summary>
public static class ClickUpService
{
	private const string ClickUpBaseUrl = "https://api.clickup.com/api/v2";

	// Business Plus plan: 1000 requests per minute per token
	private const int MaxRequestsPerMinute = 1000;

	// Retry configuration for rate limit (429) responses
	private const int MaxRetries = 3;
	private static readonly TimeSpan RateLimitWindow = TimeSpan.FromMinutes(1);
	private static readonly SemaphoreSlim RateLimitSemaphore = new(1, 1);
	private static readonly Queue<DateTime> RequestTimestamps = new();
	private static readonly TimeSpan InitialRetryDelay = TimeSpan.FromSeconds(1);

	/// <summary>
	///     Waits for a rate limit slot and records the request timestamp.
	/// </summary>
	private static async Task WaitForRateLimitAsync()
	{
		await RateLimitSemaphore.WaitAsync();
		try
		{
			var now         = DateTime.UtcNow;
			var windowStart = now - RateLimitWindow;

			// Remove timestamps outside the current window
			while (RequestTimestamps.Count > 0 && RequestTimestamps.Peek() < windowStart)
				RequestTimestamps.Dequeue();

			// If we've hit the limit, wait until the oldest request expires
			if (RequestTimestamps.Count >= MaxRequestsPerMinute)
			{
				var oldestTimestamp = RequestTimestamps.Peek();
				var waitTime        = oldestTimestamp.Add(RateLimitWindow) - now;
				if (waitTime > TimeSpan.Zero)
				{
					Debug.Print($"Rate limit reached. Waiting {waitTime.TotalSeconds:F1} seconds...");
					await Task.Delay(waitTime);
				}

				// Clean up again after waiting
				now         = DateTime.UtcNow;
				windowStart = now - RateLimitWindow;
				while (RequestTimestamps.Count > 0 && RequestTimestamps.Peek() < windowStart)
					RequestTimestamps.Dequeue();
			}

			RequestTimestamps.Enqueue(DateTime.UtcNow);
		}
		finally
		{
			RateLimitSemaphore.Release();
		}
	}

	/// <summary>
	///     Executes an HTTP request with rate limiting and retry logic for 429 responses.
	/// </summary>
	private static async Task<HttpResponseMessage> ExecuteWithRetryAsync(
		Func<Task<HttpResponseMessage>> requestFunc,
		string operationName)
	{
		var retryCount = 0;
		var retryDelay = InitialRetryDelay;

		while (true)
		{
			// Apply rate limiting before each request
			await WaitForRateLimitAsync();

			var response = await requestFunc();

			// Check for rate limit (429) response
			if ((int)response.StatusCode != 429) return response;
			if (retryCount >= MaxRetries)
			{
				Debug.Print($"{operationName}: Rate limit exceeded after {MaxRetries} retries");
				return response;
			}

			// Try to read the rate limit reset header
			var resetHeader = response.Headers.TryGetValues("X-RateLimit-Reset", out var values)
				? values.FirstOrDefault()
				: null;

			var waitTime = retryDelay;
			if (resetHeader != null && long.TryParse(resetHeader, out var resetUnixTime))
			{
				var resetTime      = DateTimeOffset.FromUnixTimeMilliseconds(resetUnixTime).UtcDateTime;
				var calculatedWait = resetTime - DateTime.UtcNow;
				if (calculatedWait > TimeSpan.Zero && calculatedWait < TimeSpan.FromMinutes(2))
					waitTime = calculatedWait;
			}

			Debug.Print(
				$"{operationName}: Rate limited (429). Retrying after {waitTime.TotalSeconds:F1} seconds... (attempt {retryCount + 1}/{MaxRetries})");
			await Task.Delay(waitTime);

			// Exponential backoff for next retry
			retryDelay = TimeSpan.FromSeconds(retryDelay.TotalSeconds * 2);
			retryCount++;
		}
	}

	/// <summary>
	///     Creates a task with subtask and checklist item for a DXF export.
	/// </summary>
	/// <param name="partNumber">The part-number to use for task naming.</param>
	/// <returns>True if all operations succeeded, false otherwise.</returns>
	public static async Task<bool> CreateDxfExportTaskAsync(string partNumber)
	{
		var options = UserOptions.Load();

		if (!options.EnableClickUpIntegration)
			return false;

		if (string.IsNullOrWhiteSpace(options.ClickUpApiToken) ||
		    string.IsNullOrWhiteSpace(options.ClickUpListId))
		{
			MessageBox.Show("ClickUp integration is enabled but API token or List ID is not configured.",
				"ClickUp Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			return false;
		}

		try
		{
			using var client = new HttpClient();
			client.DefaultRequestHeaders.Add("Authorization", options.ClickUpApiToken);
			var taskId = await CreateTaskAsync(client, options.ClickUpListId, partNumber, options.ClickUpAssigneeId);
			return taskId != null;
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error creating ClickUp task: {ex.Message}",
				"ClickUp Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			return false;
		}
	}

	/// <summary>
	///     Searches for an open task with the specified name in the given list.
	/// </summary>
	private static async Task<string?> FindOpenTaskByNameAsync(HttpClient client, string listId, string taskName)
	{
		var url =
			$"{ClickUpBaseUrl}/list/{listId}/task?name={Uri.EscapeDataString(taskName)}&statuses[]=open&statuses[]=to%20do&statuses[]=in%20progress";

		var response = await ExecuteWithRetryAsync(() => client.GetAsync(url), "FindOpenTask");
		if (!response.IsSuccessStatusCode)
		{
			var error = await response.Content.ReadAsStringAsync();
			Debug.Print($"Failed to search for existing task: {error}");
			return null;
		}

		var responseJson = await response.Content.ReadAsStringAsync();
		var doc          = JsonDocument.Parse(responseJson);

		if (!doc.RootElement.TryGetProperty("tasks", out var tasks) || tasks.GetArrayLength() == 0)
			return null;

		foreach (var task in tasks.EnumerateArray())
		{
			if (!task.TryGetProperty("name", out var nameElement) ||
			    nameElement.GetString()?.Equals(taskName, StringComparison.OrdinalIgnoreCase) != true) continue;
			if (task.TryGetProperty("id", out var idElement))
				return idElement.GetString();
		}

		return null;
	}

	/// <summary>
	///     Creates a task in the specified list. If an open task with the same name already exists, returns the existing task
	///     ID.
	/// </summary>
	private static async Task<string?> CreateTaskAsync(HttpClient client, string listId, string partNumber,
		string? assigneeId)
	{
		var existingTaskId = await FindOpenTaskByNameAsync(client, listId, partNumber);
		if (existingTaskId != null)
		{
			Debug.Print($"Task with name '{partNumber}' already exists (ID: {existingTaskId}). Skipping creation.");
			return existingTaskId;
		}

		var url = $"{ClickUpBaseUrl}/list/{listId}/task";

		var taskData = new Dictionary<string, object>
		{
			["name"] = $"{partNumber}",
			["description"] =
				$"DXF file exported for part number: {partNumber}\n\nExport completed at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
			["status"] = "TO DO"
		};

		if (!string.IsNullOrWhiteSpace(assigneeId)) taskData["assignees"] = new[] { int.Parse(assigneeId) };

		var json    = JsonSerializer.Serialize(taskData);
		var content = new StringContent(json, Encoding.UTF8, "application/json");

		var response = await ExecuteWithRetryAsync(() => client.PostAsync(url, content), "CreateTask");
		if (!response.IsSuccessStatusCode)
		{
			var error = await response.Content.ReadAsStringAsync();
			Debug.Print($"Failed to create task: {error}");
			return null;
		}

		var responseJson = await response.Content.ReadAsStringAsync();
		var doc          = JsonDocument.Parse(responseJson);
		return doc.RootElement.GetProperty("id").GetString();
	}

	/// <summary>
	///     Validates ClickUp configuration by making a test API call.
	/// </summary>
	public static async Task<bool> ValidateConfigurationAsync(string token)
	{
		try
		{
			using var client = new HttpClient();
			client.DefaultRequestHeaders.Add("Authorization", token);

			var response =
				await ExecuteWithRetryAsync(() => client.GetAsync($"{ClickUpBaseUrl}/user"), "ValidateConfig");
			return response.IsSuccessStatusCode;
		}
		catch
		{
			return false;
		}
	}
}