#nullable enable
namespace DoyleAddin.Optional_Features.BatchExport;

using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using DXFs;
using Inventor;
using Svg;
using Path = Path;
using ThemeManager = Options.Themes.ThemeManager;

public partial class BatchExportForm
{
	private AssemblyDocument? _assemblyDoc;

	public BatchExportForm()
	{
		InitializeComponent();
		Loaded += BatchExportForm_Loaded;
	}

	public static event EventHandler? RequestClose;

	private void BatchExportForm_Loaded(object sender, RoutedEventArgs e)
	{
		try
		{
			ThemeManager.ApplyTheme(this);

			var partIcon     = TryLoadSvgIcon("DoyleAddin.Resources.Icons.PartIcon.svg");
			var assemblyIcon = TryLoadSvgIcon("DoyleAddin.Resources.Icons.AssemblyIcon.svg");

			_assemblyDoc = (AssemblyDocument)ThisApplication.ActiveDocument;
			var seenParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
			PopulateTreeView(_assemblyDoc.ComponentDefinition.Occurrences, treeViewParts.Items, partIcon,
				assemblyIcon, seenParts);
		}
		catch (Exception ex)
		{
			MessageBox.Show("Error initializing form: " + ex.Message, "Error", MessageBoxButton.OK,
				MessageBoxImage.Error);
			RequestClose?.Invoke(this, EventArgs.Empty);
		}
	}

	private static BitmapSource? TryLoadSvgIcon(string resourceName)
	{
		try
		{
			using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);
			if (stream == null) return null;

			var svgDoc = SvgDocument.Open<SvgDocument>(stream);
			var layer = svgDoc.Descendants().OfType<SvgGroup>().FirstOrDefault(g =>
				(g.CustomAttributes.ContainsKey("inkscape:label") &&
				 g.CustomAttributes["inkscape:label"] == "Dark") ||
				(g.CustomAttributes.ContainsKey("http://www.inkscape.org/namespaces/inkscape:label") &&
				 g.CustomAttributes["http://www.inkscape.org/namespaces/inkscape:label"] == "Dark"));

			SvgDocument docToRender;
			if (layer != null)
			{
				docToRender = new SvgDocument { Width = svgDoc.Width, Height = svgDoc.Height };
				docToRender.Children.Add(layer.DeepCopy());
			}
			else
			{
				docToRender = svgDoc;
			}

			using var bitmap = docToRender.Draw(16, 16);
			if (bitmap == null) return null;

			var hBitmap = bitmap.GetHbitmap();
			try
			{
				return Imaging.CreateBitmapSourceFromHBitmap(hBitmap, IntPtr.Zero, Int32Rect.Empty,
					BitmapSizeOptions.FromEmptyOptions());
			}
			finally
			{
				DeleteObject(hBitmap);
			}
		}
		catch
		{
			return null;
		}
	}

	[LibraryImport("gdi32.dll")]
	private static partial void DeleteObject(IntPtr hObject);

	private static TreeViewItem? FindParentTreeViewItem(DependencyObject? element)
	{
		while (element != null)
		{
			if (element is TreeViewItem item)
				return item;
			element = VisualTreeHelper.GetParent(element);
		}

		return null;
	}

	private static CheckBox? FindParentCheckBox(DependencyObject? element)
	{
		while (element != null)
		{
			if (element is CheckBox cb)
				return cb;
			element = VisualTreeHelper.GetParent(element);
		}

		return null;
	}

	private static bool IsClickOnExpandCollapseButton(DependencyObject? element)
	{
		while (element != null)
		{
			if (element.GetType().Name == "ToggleButton")
				return true;
			if (element is TreeViewItem)
				return false;
			element = VisualTreeHelper.GetParent(element);
		}

		return false;
	}

	/// <summary>Populates <paramref name="nodes" /> and returns true if any sheet metal parts were added.</summary>
	private bool PopulateTreeView(ComponentOccurrences occurrences,
		ItemCollection nodes, ImageSource? partIcon, ImageSource? assemblyIcon,
		HashSet<string> seenParts)
	{
		var addedAny = false;

		foreach (ComponentOccurrence occurrence in occurrences)
			try
			{
				switch (occurrence.DefinitionDocumentType)
				{
					case kPartDocumentObject:
					{
						var partDoc = (PartDocument)occurrence.Definition.Document;
						if (partDoc.ComponentDefinition is not SheetMetalComponentDefinition)
							break;

						if (!seenParts.Add(partDoc.FullFileName))
							break;

						var checkBox = new CheckBox
						{
							Content           = Path.GetFileNameWithoutExtension(partDoc.FullFileName),
							IsChecked         = false,
							Tag               = partDoc,
							VerticalAlignment = VerticalAlignment.Center
						};

						var treeItem = new TreeViewItem
						{
							Header  = CreateHeaderStackPanel(partIcon, checkBox),
							Tag     = partDoc,
							Padding = new Thickness(2, 1, 2, 1)
						};

						// Clicking anywhere else on the item selects it in Inventor
						treeItem.PreviewMouseLeftButtonDown += (_, e) =>
						{
							if (FindParentCheckBox(e.OriginalSource as DependencyObject) != null)
								return; // Let checkbox handle its own click

							if (IsClickOnExpandCollapseButton(e.OriginalSource as DependencyObject))
								return; // Let expand/collapse button work

							// Only handle if click is on this item, not a child
							var clickedItem = FindParentTreeViewItem(e.OriginalSource as DependencyObject);
							if (clickedItem != treeItem)
								return;

							treeItem.IsSelected = true;
							e.Handled           = true;
						};
						checkBox.Checked       += (_, _) => UpdateParentCheckboxesUpTheTree(treeItem);
						checkBox.Unchecked     += (_, _) => UpdateParentCheckboxesUpTheTree(treeItem);
						checkBox.Indeterminate += (_, _) => UpdateParentCheckboxesUpTheTree(treeItem);

						treeItem.Selected += (_, args) =>
						{
							SelectOccurrencesInInventor(partDoc);
							args.Handled = true;
						};

						var convertMenuItem = new MenuItem { Header = "Convert to Standard Part" };
						convertMenuItem.Click += (_, _) => ConvertToStandardPart(partDoc, treeItem);
						treeItem.ContextMenu  =  new ContextMenu { Items = { convertMenuItem } };

						nodes.Add(treeItem);
						addedAny = true;
						break;
					}

					case kAssemblyDocumentObject:
					{
						var assemblyNode = new TreeViewItem
						{
							IsExpanded = true,
							Padding    = new Thickness(2, 1, 2, 1)
						};

						var hasChildren = PopulateTreeView(occurrence.Definition.Occurrences,
							assemblyNode.Items, partIcon, assemblyIcon, seenParts);

						if (!hasChildren)
							break;

						var assemblyCheckBox = new CheckBox
						{
							Content           = occurrence.Name,
							IsChecked         = false,
							VerticalAlignment = VerticalAlignment.Center,
							IsThreeState      = true // Enables indeterminate state
						};

						assemblyCheckBox.Checked   += (_, _) => SetAllChildCheckboxes(assemblyNode.Items, true);
						assemblyCheckBox.Unchecked += (_, _) => SetAllChildCheckboxes(assemblyNode.Items, false);
						assemblyNode.Header        =  CreateHeaderStackPanel(assemblyIcon, assemblyCheckBox);

						var capturedOccurrence = occurrence;
						assemblyNode.PreviewMouseLeftButtonDown += (_, e) =>
						{
							if (FindParentCheckBox(e.OriginalSource as DependencyObject) != null)
								return;

							if (IsClickOnExpandCollapseButton(e.OriginalSource as DependencyObject))
								return; // Let expand/collapse button work

							// Only handle if click is on this item, not a child
							var clickedItem = FindParentTreeViewItem(e.OriginalSource as DependencyObject);
							if (clickedItem != assemblyNode)
								return;

							assemblyNode.IsSelected = true;
							e.Handled               = true;
						};

						assemblyNode.Selected += (_, args) =>
						{
							SelectSingleOccurrenceInInventor(capturedOccurrence);
							args.Handled = true;
						};

						nodes.Add(assemblyNode);
						addedAny = true;
						break;
					}
				}
			}
			catch
			{
				// Skip occurrences that throw COM errors
			}

		return addedAny;
	}

	private static void UpdateParentCheckboxesUpTheTree(TreeViewItem item)
	{
		var parent = item.Parent as TreeViewItem;
		while (parent != null)
		{
			UpdateParentCheckbox(parent);
			parent = parent.Parent as TreeViewItem;
		}
	}

	private static StackPanel CreateHeaderStackPanel(ImageSource? icon, CheckBox checkBox)
	{
		var sp = new StackPanel { Orientation = Orientation.Horizontal };

		if (icon != null)
			sp.Children.Add(new Image
			{
				Source            = icon,
				Width             = 16,
				Height            = 16,
				Margin            = new Thickness(0, 0, 4, 0),
				VerticalAlignment = VerticalAlignment.Center
			});

		sp.Children.Add(checkBox);
		return sp;
	}

	/// <summary>
	///     Updates an assembly checkbox to reflect the checked state of its children (true / false / indeterminate).
	///     Call this after any child checkbox changes.
	/// </summary>
	private static void UpdateParentCheckbox(TreeViewItem? parent)
	{
		if (parent?.Header is not StackPanel sp) return;

		var parentCb = sp.Children.OfType<CheckBox>().FirstOrDefault();
		if (parentCb == null) return;

		var checkedCount = 0;
		var total        = 0;

		foreach (TreeViewItem child in parent.Items)
			if (child.Header is StackPanel childSp)
			{
				var childCb = childSp.Children.OfType<CheckBox>().FirstOrDefault();
				if (childCb?.IsChecked == true) checkedCount++;
				total++;
			}

		var newState = checkedCount == 0 ? false :
			checkedCount == total        ? true : (bool?)null;

		parentCb.SetCurrentValue(ToggleButton.IsCheckedProperty, newState);
	}

	/// <summary>
	///     Selects all occurrences of <paramref name="partDoc" /> in the active Inventor assembly.
	/// </summary>
	private void SelectOccurrencesInInventor(PartDocument partDoc)
	{
		try
		{
			if (_assemblyDoc == null) return;

			var selectSet = _assemblyDoc.SelectSet;
			selectSet.Clear();

			var occurrences = new List<ComponentOccurrence>();
			CollectOccurrencesByPart(_assemblyDoc.ComponentDefinition.Occurrences,
				partDoc.FullFileName, occurrences);

			foreach (var occ in occurrences)
				selectSet.Select(occ);
		}
		catch
		{
			// Swallow COM/interop errors silently — selection is a convenience feature.
		}
	}

	/// <summary>
	///     Selects a single <paramref name="occurrence" /> in the active Inventor assembly.
	/// </summary>
	private void SelectSingleOccurrenceInInventor(ComponentOccurrence occurrence)
	{
		try
		{
			if (_assemblyDoc == null) return;

			var selectSet = _assemblyDoc.SelectSet;
			selectSet.Clear();
			selectSet.Select(occurrence);
		}
		catch
		{
			// Swallow COM/interop errors silently.
		}
	}

	/// <summary>
	///     Converts <paramref name="partDoc" /> to a standard (non-sheet-metal) part in Inventor,
	///     then removes its tree node and prunes any assembly parents that are now empty.
	/// </summary>
	private void ConvertToStandardPart(PartDocument partDoc, TreeViewItem item)
	{
		// Setting SubType to the standard part GUID triggers Inventor's conversion.
		const string standardPartSubType = "{4D29B490-49B2-11D0-93C3-7E0706000000}";

		try
		{
			var filePath = partDoc.FullFileName;
			partDoc.Close();
			var openedDoc = (PartDocument)ThisApplication.Documents.Open(filePath);
			openedDoc.SubType = standardPartSubType;
			openedDoc.Save2(false);
			openedDoc.Close();
		}
		catch (Exception ex)
		{
			MessageBox.Show("Could not convert part: " + ex.Message, "Convert Failed",
				MessageBoxButton.OK, MessageBoxImage.Warning);
			return;
		}

		// Remove the node from whichever collection owns it, then prune empty assemblies.
		RemoveFromParent(treeViewParts.Items, item);
		PruneEmptyAssemblyNodes(treeViewParts.Items);
	}

	/// <summary>Removes <paramref name="target" /> from the tree by searching recursively.</summary>
	private static bool RemoveFromParent(ItemCollection nodes, TreeViewItem target)
	{
		for (var i = 0; i < nodes.Count; i++)
		{
			if (nodes[i] is not TreeViewItem node) continue;
			if (ReferenceEquals(node, target))
			{
				nodes.RemoveAt(i);
				return true;
			}

			if (RemoveFromParent(node.Items, target)) return true;
		}

		return false;
	}

	/// <summary>
	///     Recursively removes assembly nodes (no <see cref="PartDocument" /> Tag) that have
	///     no remaining children, bottom-up.
	/// </summary>
	private static void PruneEmptyAssemblyNodes(ItemCollection nodes)
	{
		for (var i = nodes.Count - 1; i >= 0; i--)
		{
			if (nodes[i] is not TreeViewItem node) continue;
			if (node.HasItems) PruneEmptyAssemblyNodes(node.Items);
			if (node.Items.Count == 0 && node.Tag is not PartDocument)
				nodes.RemoveAt(i);
		}
	}

	/// <summary>
	///     Recursively collects all <see cref="ComponentOccurrence" /> items whose
	///     definition document matches <paramref name="fullFileName" />.
	/// </summary>
	private static void CollectOccurrencesByPart(IEnumerable occurrences,
		string fullFileName, List<ComponentOccurrence> result)
	{
		foreach (var occ in occurrences.Cast<ComponentOccurrence>())
			try
			{
				switch (occ.DefinitionDocumentType)
				{
					case kPartDocumentObject:
						if (string.Equals(
							    ((PartDocument)occ.Definition.Document).FullFileName,
							    fullFileName,
							    StringComparison.OrdinalIgnoreCase))
							result.Add(occ);
						break;
					case kAssemblyDocumentObject:
						CollectOccurrencesByPart(occ.SubOccurrences, fullFileName, result);
						break;
				}
			}
			catch
			{
				// Skip unresolved/suppressed occurrences.
			}
	}

	private static void SetAllChildCheckboxes(ItemCollection nodes, bool isChecked)
	{
		foreach (TreeViewItem node in nodes)
		{
			if (node.Header is StackPanel sp)
			{
				var cb = sp.Children.OfType<CheckBox>().FirstOrDefault();
				cb?.SetCurrentValue(ToggleButton.IsCheckedProperty, (bool?)isChecked);
			}

			if (node.HasItems)
				SetAllChildCheckboxes(node.Items, isChecked);
		}
	}

	private void buttonSelectAll_Click(object sender, RoutedEventArgs e)
	{
		SetAllChildCheckboxes(treeViewParts.Items, true);
	}

	private void buttonExport_Click(object sender, RoutedEventArgs e)
	{
		var selectedParts = new List<PartDocument>();
		GetCheckedParts(treeViewParts.Items, selectedParts);

		if (selectedParts.Count == 0)
		{
			MessageBox.Show("No sheet metal parts selected for export.", "No Selection", MessageBoxButton.OK,
				MessageBoxImage.Information);
			return;
		}

		DxfUpdate.BatchExport(selectedParts);
		RequestClose?.Invoke(this, EventArgs.Empty);
	}

	private void buttonCancel_Click(object sender, RoutedEventArgs e)
	{
		RequestClose?.Invoke(this, EventArgs.Empty);
	}

	private static void GetCheckedParts(ItemCollection nodes, List<PartDocument> selectedParts)
	{
		foreach (TreeViewItem node in nodes)
		{
			if (node.Header is StackPanel stackPanel)
			{
				var checkBox = stackPanel.Children.OfType<CheckBox>().FirstOrDefault();
				if (checkBox is { IsChecked: true, Tag: PartDocument partDoc })
					selectedParts.Add(partDoc);
			}

			if (node.HasItems) GetCheckedParts(node.Items, selectedParts);
		}
	}
}