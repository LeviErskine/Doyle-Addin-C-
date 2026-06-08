namespace DoyleAddin.My_Project;

using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using Inventor;
using MessageBox = System.Windows.Forms.MessageBox;

/// <summary>
///     Interface for panel wrapper to break bidirectional dependency.
/// </summary>
public interface IPanelWrapper
{
	/// <summary>
	/// </summary>
	bool IsDisposed { get; }

	/// <summary>
	///     Closes the panel wrapper.
	/// </summary>
	void Close();
}

/// <summary>
///     Adapter for properly hosting WPF windows in Inventor dockable windows with keyboard support.
/// </summary>
/// <remarks>
///     Adapter for properly hosting WPF windows in Inventor dockable windows with keyboard support.
/// </remarks>
public class DockableWindowChildAdapter
{
	private const uint DLGC_WANTARROWS = 0x1;
	private const uint DLGC_WANTTAB = 0x2;
	private const uint DLGC_WANTALLKEYS = 0x4;
	private const uint DLGC_HASSETSEL = 0x8;
	private const uint DLGC_WANTCHARS = 0x80;
	private const uint WM_GETDLGCODE = 0x87;
	private readonly DockableWindow _dockableWindow;
	private readonly IPanelWrapper _panelWrapper;
	private HwndSource _hwndSource;
	private DispatcherTimer _visibilityTimer;

	/// <summary>
	///     Adapter for properly hosting WPF windows in Inventor dockable windows with keyboard support.
	/// </summary>
	/// <remarks>
	///     This class provides functionality to integrate and manage WPF windows
	///     within Inventor's dockable window framework, ensuring proper handling
	///     of keyboard shortcuts and resources.
	/// </remarks>
	public DockableWindowChildAdapter(DockableWindow dockableWindow, IPanelWrapper panelWrapper)
	{
		ArgumentNullException.ThrowIfNull(dockableWindow);
		ArgumentNullException.ThrowIfNull(panelWrapper);

		_dockableWindow = dockableWindow;
		_panelWrapper   = panelWrapper;
	}

	/// <summary>
	///     Adds a WPF window to an Inventor dockable window with proper keyboard input handling.
	/// </summary>
	/// <param name="window">The WPF window to host.</param>
	/// <exception cref="ArgumentNullException">Thrown when the window is null.</exception>
	public void AddWPFWindow(Window window)
	{
		ArgumentNullException.ThrowIfNull(window);

		window.WindowStyle = WindowStyle.None;
		window.WindowState = WindowState.Maximized;
		window.ResizeMode  = ResizeMode.NoResize;
		window.Show();

		var win = Window.GetWindow(window);
		if (win == null) return;
		var wih  = new WindowInteropHelper(win);
		var hWnd = wih.EnsureHandle();

		_dockableWindow.AddChild(hWnd);

		_hwndSource = HwndSource.FromHwnd(hWnd);
		_hwndSource?.AddHook(WndProc);

		// Start timer to monitor dockable window visibility
		_visibilityTimer = new DispatcherTimer
		{
			Interval = TimeSpan.FromMilliseconds(500)
		};
		_visibilityTimer.Tick += (_, _) =>
		{
			if (_dockableWindow.Visible && !_panelWrapper.IsDisposed) return;
			_visibilityTimer.Stop();
			_panelWrapper.Close();
		};
		_visibilityTimer.Start();
	}

	/// <summary>
	///     Cleans up timer and hook resources during disposal.
	/// </summary>
	public void Cleanup()
	{
		_visibilityTimer?.Stop();
		_hwndSource?.RemoveHook(WndProc);
	}

	/// <summary>
	///     Window procedure hook for handling keyboard input messages.
	/// </summary>
	private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
	{
		if (msg == WM_GETDLGCODE)
		{
			handled = true;
			return new IntPtr(DLGC_WANTCHARS | DLGC_WANTARROWS | DLGC_HASSETSEL | DLGC_WANTTAB | DLGC_WANTALLKEYS);
		}

		// Check for close message (WM_CLOSE = 0x0010)
		if (msg != 0x0010) return IntPtr.Zero;

		// Guard against recursive close calls during disposal
		if (_panelWrapper?.IsDisposed == true) return IntPtr.Zero;

		handled = true;
		_panelWrapper?.Close();
		return IntPtr.Zero;
	}
}

/// <summary>
///     Wraps a WPF UIElement into an Inventor dockable window.
/// </summary>
[ComVisible(true)]
public class PanelWrapper : IDisposable, IPanelWrapper
{
	private readonly UIElement _content;
	private DockableWindowChildAdapter _adapter;
	private DockableWindow _dockableWindow;
	private Window _wpfWindow;

	/// <summary>
	///     Initializes a new instance of the PanelWrapper class.
	/// </summary>
	/// <param name="content">The WPF UIElement to host in the dockable window.</param>
	/// <param name="title">The title for the dockable window.</param>
	public PanelWrapper(UIElement content, string title)
	{
		ArgumentNullException.ThrowIfNull(content);
		ArgumentNullException.ThrowIfNull(title);

		_content = content;
		CreateDockableWindow(title);
	}

	/// <summary>
	///     Releases all resources used by the PanelWrapper.
	/// </summary>
	public void Dispose()
	{
		Dispose(true);
		GC.SuppressFinalize(this);
	}

	/// <summary>
	/// </summary>
	public bool IsDisposed { get; private set; }

	/// <summary>
	///     Closes the dockable window and properly cleans up resources.
	/// </summary>
	public void Close()
	{
		Dispose();
	}

	/// <summary>
	///     Creates the dockable window and sets its content.
	/// </summary>
	private void CreateDockableWindow(string title)
	{
		if (ThisApplication?.UserInterfaceManager == null)
		{
			MessageBox.Show("Invalid Inventor application instance.");
			return;
		}

		try
		{
			// Create a WPF window to host the control
			_wpfWindow = new Window
			{
				Content               = _content,
				WindowStartupLocation = WindowStartupLocation.CenterScreen
			};

			// Hook up the closing event to handle built-in close button
			_wpfWindow.Closing += (_, args) =>
			{
				// During disposal, allow normal close; otherwise prevent it
				if (IsDisposed) return;
				args.Cancel = true;
				Close();
			};

			// Create the Inventor dockable window
			_dockableWindow = ThisApplication.UserInterfaceManager.DockableWindows.Add(
				"DoyleAddin.My_Project.PanelWrapper", // Internal name
				title,                                // Title
				Globals.AddInClientId());             // Client ID

			// Create adapter instance and use it to host the WPF window
			_adapter = new DockableWindowChildAdapter(_dockableWindow, this);
			_adapter.AddWPFWindow(_wpfWindow);
			_dockableWindow.ShowTitleBar = true;
			_dockableWindow.Visible      = true;
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error creating dockable window: {ex.Message}");
		}
	}

	/// <summary>
	///     Shows the dockable window.
	/// </summary>
	public void Show()
	{
		_dockableWindow?.Visible = true;
	}

	/// <summary>
	///     Hides the dockable window.
	/// </summary>
	public void Hide()
	{
		_dockableWindow?.Visible = false;
	}

	/// <summary>
	///     Releases the unmanaged resources and optionally releases the managed resources.
	/// </summary>
	/// <param name="disposing">
	///     true to release both managed and unmanaged resources; false to release only unmanaged
	///     resources.
	/// </param>
	protected virtual void Dispose(bool disposing)
	{
		if (IsDisposed || !disposing) return;

		IsDisposed = true;

		try
		{
			if (_dockableWindow != null)
			{
				try
				{
					_dockableWindow.Clear();
				}
				catch (Exception)
				{
					// Ignore errors during cleanup
				}

				_dockableWindow.Visible = false;

				try
				{
					_dockableWindow.Delete();
				}
				catch (Exception)
				{
					// Ignore errors during cleanup
				}

				_dockableWindow = null;
			}

			if (_wpfWindow != null)
			{
				try
				{
					_wpfWindow.Close();
				}
				catch (Exception)
				{
					// Ignore errors during cleanup
				}

				_wpfWindow = null;
			}

			if (_adapter == null) return;
			_adapter.Cleanup();
			_adapter = null;
		}
		catch (Exception)
		{
			// Ignore errors during disposal
		}
	}
}