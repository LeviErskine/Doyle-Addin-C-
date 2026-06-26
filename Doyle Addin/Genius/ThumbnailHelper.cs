namespace DoyleAddin.Genius;

using System.Windows.Media.Imaging;
using OpenMcdf;

public static class ThumbnailHelper
{
	public static BitmapSource GetThumbnail(Document document)
	{
		if (document == null)
		{
			Debug.WriteLine("ThumbnailHelper.GetThumbnail: document is null");
			return null;
		}

		string filePath;
		try
		{
			filePath = document.FullFileName;
		}
		catch
		{
			filePath = null;
		}

		if (string.IsNullOrEmpty(filePath) || !Path.Exists(filePath))
		{
			Debug.WriteLine($"ThumbnailHelper.GetThumbnail: no file path for '{document.DisplayName}'");
			return null;
		}

		try
		{
			return ExtractThumbnailFromFile(filePath);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"ThumbnailHelper.GetThumbnail: file extraction failed: {ex.Message}");
			return null;
		}
	}

	private static BitmapFrame ExtractThumbnailFromFile(string filePath)
	{
		using var root = RootStorage.OpenRead(filePath);

		foreach (var entry in root.EnumerateEntries())
		{
			if (entry.Name is not { Length: > 0 } || entry.Name[0] != '\u0005') continue;
			if (!root.TryOpenStream(entry.Name, out var stream) || stream.Length < 100)
			{
				stream?.Dispose();
				continue;
			}

			byte[] raw;
			using (stream)
			using (var ms = new MemoryStream((int)stream.Length))
			{
				stream.CopyTo(ms);
				raw = ms.ToArray();
			}

			var pngOffset = IndexOfPattern(raw, [0x89, 0x50, 0x4E, 0x47]);
			if (pngOffset < 0) continue;

			using var pngStream = new MemoryStream(raw, pngOffset, raw.Length - pngOffset);
			var       decoder   = BitmapDecoder.Create(pngStream, BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
			var       frame     = decoder.Frames[0];
			if (frame.CanFreeze) frame.Freeze();
			return frame;
		}

		return null;
	}

	private static int IndexOfPattern(byte[] data, byte[] pattern)
	{
		if (data.Length < pattern.Length) return -1;
		for (var i = 0; i <= data.Length - pattern.Length; i++)
		{
			var match = !pattern.Where((t, j) => data[i + j] != t).Any();

			if (match) return i;
		}

		return -1;
	}
}