<#
    'EP PDF TOOLS' is application based on powershell for pdf merge and
    split using iTextSharp assembly for PDFs manipulation.

    Copyright (C) 2017  Vladimir Mihhejenko

    This program is free software: you can redistribute it and/or modify 
    it under the terms of the GNU Affero General Public License 
    as published by the Free Software Foundation, either version 3 of the 
    License, or any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU Affero General Public License for more details.

    You should have received a copy of the GNU Affero General Public License 
    along with this program. If not, see <http://www.gnu.org/licenses/>. Or
    write to the Free Software Foundation, Inc., 51 Franklin Street, 
    Fifth Floor, Boston, MA 02110-1301 USA.

    For contacts use vovit@mail.ru
#>
# ASSEMBLIES
Add-Type -AssemblyName PresentationFramework
Add-Type -Path $PSScriptRoot\itextsharp.dll # iTextSharp library path

# GET-FOLDER FUNCTION
Function Get-Folder {
$sourcecode = @"
using System;
using System.Windows.Forms;
using System.Reflection;
namespace FolderSelect
{
	public class FolderSelectDialog
	{
		System.Windows.Forms.OpenFileDialog ofd = null;
		public FolderSelectDialog()
		{
			ofd = new System.Windows.Forms.OpenFileDialog();
			//ofd.Filter = "Folders|\n";
			ofd.AddExtension = false;
			ofd.CheckFileExists = false;
			ofd.DereferenceLinks = true;
			ofd.Multiselect = false;
		}
		public string InitialDirectory
		{
			get { return ofd.InitialDirectory; }
			set { ofd.InitialDirectory = value == null || value.Length == 0 ? Environment.CurrentDirectory : value; }
		}
		public string Title
		{
			get { return ofd.Title; }
			set { ofd.Title = value == null ? "Select a folder" : value; }
		}
		public string FileName
		{
			get { return ofd.FileName; }
		}
		public bool ShowDialog()
		{
			return ShowDialog(IntPtr.Zero);
		}
		public bool ShowDialog(IntPtr hWndOwner)
		{
			bool flag = false;

			if (Environment.OSVersion.Version.Major >= 6)
			{
				var r = new Reflector("System.Windows.Forms");
				uint num = 0;
				Type typeIFileDialog = r.GetType("FileDialogNative.IFileDialog");
				object dialog = r.Call(ofd, "CreateVistaDialog");
				r.Call(ofd, "OnBeforeVistaDialog", dialog);
				uint options = (uint)r.CallAs(typeof(System.Windows.Forms.FileDialog), ofd, "GetOptions");
				options |= (uint)r.GetEnum("FileDialogNative.FOS", "FOS_PICKFOLDERS");
				r.CallAs(typeIFileDialog, dialog, "SetOptions", options);
				object pfde = r.New("FileDialog.VistaDialogEvents", ofd);
				object[] parameters = new object[] { pfde, num };
				r.CallAs2(typeIFileDialog, dialog, "Advise", parameters);
				num = (uint)parameters[1];
				try
				{
					int num2 = (int)r.CallAs(typeIFileDialog, dialog, "Show", hWndOwner);
					flag = 0 == num2;
				}
				finally
				{
					r.CallAs(typeIFileDialog, dialog, "Unadvise", num);
					GC.KeepAlive(pfde);
				}
			}
			else
			{
				var fbd = new FolderBrowserDialog();
				fbd.Description = this.Title;
				fbd.SelectedPath = this.InitialDirectory;
				fbd.ShowNewFolderButton = false;
				if (fbd.ShowDialog(new WindowWrapper(hWndOwner)) != DialogResult.OK) return false;
				ofd.FileName = fbd.SelectedPath;
				flag = true;
			}
			return flag;
		}
	}
	public class WindowWrapper : System.Windows.Forms.IWin32Window
	{
		public WindowWrapper(IntPtr handle)
		{
			_hwnd = handle;
		}
		public IntPtr Handle
		{
			get { return _hwnd; }
		}

		private IntPtr _hwnd;
	}
	public class Reflector
	{
		string m_ns;
		Assembly m_asmb;
		public Reflector(string ns)
			: this(ns, ns)
		{ }
		public Reflector(string an, string ns)
		{
			m_ns = ns;
			m_asmb = null;
			foreach (AssemblyName aN in Assembly.GetExecutingAssembly().GetReferencedAssemblies())
			{
				if (aN.FullName.StartsWith(an))
				{
					m_asmb = Assembly.Load(aN);
					break;
				}
			}
		}
		public Type GetType(string typeName)
		{
			Type type = null;
			string[] names = typeName.Split('.');

			if (names.Length > 0)
				type = m_asmb.GetType(m_ns + "." + names[0]);

			for (int i = 1; i < names.Length; ++i) {
				type = type.GetNestedType(names[i], BindingFlags.NonPublic);
			}
			return type;
		}
		public object New(string name, params object[] parameters)
		{
			Type type = GetType(name);
			ConstructorInfo[] ctorInfos = type.GetConstructors();
			foreach (ConstructorInfo ci in ctorInfos) {
				try {
					return ci.Invoke(parameters);
				} catch { }
			}

			return null;
		}
		public object Call(object obj, string func, params object[] parameters)
		{
			return Call2(obj, func, parameters);
		}
		public object Call2(object obj, string func, object[] parameters)
		{
			return CallAs2(obj.GetType(), obj, func, parameters);
		}
		public object CallAs(Type type, object obj, string func, params object[] parameters)
		{
			return CallAs2(type, obj, func, parameters);
		}
		public object CallAs2(Type type, object obj, string func, object[] parameters) {
			MethodInfo methInfo = type.GetMethod(func, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
			return methInfo.Invoke(obj, parameters);
		}
		public object Get(object obj, string prop)
		{
			return GetAs(obj.GetType(), obj, prop);
		}
		public object GetAs(Type type, object obj, string prop) {
			PropertyInfo propInfo = type.GetProperty(prop, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
			return propInfo.GetValue(obj, null);
		}
		public object GetEnum(string typeName, string name) {
			Type type = GetType(typeName);
			FieldInfo fieldInfo = type.GetField(name);
			return fieldInfo.GetValue(null);
		}
	}
}
"@

$assemblies = ('System.Windows.Forms', 'System.Reflection')
Add-Type -TypeDefinition $sourceCode -ReferencedAssemblies $assemblies -ErrorAction STOP
$fsd = New-Object FolderSelect.FolderSelectDialog
    $fsd.Title = "What to select"
    $fsd.ShowDialog() | Out-Null
    $returnFolder = $fsd.FileName + '\'
return $returnFolder }

Function PDF-Merge-Process {

    Param( $inputFolder, [String]$fileName, [String]$filter )

    $document = New-Object iTextSharp.text.Document
    $fileStream = New-Object System.IO.FileStream($fileName, [System.IO.FileMode]::Create)
    $pdfCopy = New-Object iTextSharp.text.pdf.PdfCopy($document, $fileStream)
    $document.Open()

    Get-ChildItem -Path $inputFolder -Filter $filter | ForEach-Object {
        $reader = New-Object iTextSharp.text.pdf.PdfReader($_.FullName)
        $pdfCopy.AddDocument($reader)
        $reader.Dispose() }
    $pdfCopy.Dispose()
    $fileStream.Dispose()
    $document.Dispose()
}

Function Merge-PDF {            

    Param( [String]$inputFolder, [String]$outputFolder )

  if((Test-Path $inputFolder) -and (Test-Path $outputFolder) -and # check paths exists
   $inputFolder -ne '\' -and $outputFolder -ne '\' ) { # and not \
    if (! $outputFolder.EndsWith('\')) { $outputFolder += '\' }
    if (! $inputFolder.EndsWith('\')) { $inputFolder += '\' }
    $splitedFolders = Get-ChildItem -Path $inputFolder -Filter *.pdf -Directory
    $timeStamp = 'Merged_'+ (Get-Date -f HHmmss-dd-MMMM-yyyy)
    $newFolderName = $outputFolder + $timeStamp 
    New-Item  $newFolderName -Type Directory
    if ($splitedFolders.Count -ge 1) { 
        foreach ($folder in $splitedFolders) {
             $fileFullName = $folder.FullName + '\'
             $fileName = $folder.FullName.Split('\') | select -last 1
             if (! $newFolderName.EndsWith('\')) { $newFolderName += '\' }
             $mergedOutput = $newFolderName + $fileName
             PDF-Merge-Process -inputFolder $fileFullName -fileName $mergedOutput -filter *.pdf
             if ( $fileFullName.EndsWith('\') ) { $fileFullName.Trim('\') }
             Remove-Item $fileFullName -Recurse -force -ErrorAction SilentlyContinue -Confirm:$false
        } } #end if
     else { $mergedFilename = $newFolderName + '\' + $timeStamp + '.pdf'
            PDF-Merge-Process -inputFolder $inputFolder -fileName $mergedFilename -filter *.pdf
            }      
     } <# end of IF newFolder exists #>   
}

Function Split-PDF {
    
Param( [String]$inputFolder, [String]$outputFolder )

if( (Test-Path $inputFolder) -and (Test-Path $outputFolder) -and # check paths exists
   $inputFolder -ne '\' -and $outputFolder -ne '\' -and # and not \
   $inputFolder -ne $outputFolder ) { # and notequal each other)  
    # get each PDF file in inputFolder directory
    if(! $outputFolder.EndsWith('\')) { $outputFolder += '\' }
    Get-ChildItem -Path $inputFolder -Filter *.pdf | ForEach-Object {
      $reader = New-Object iTextSharp.text.pdf.PdfReader($_.FullName) # read file 
      if ($reader.NumberOfPages -ne 1 ) {
      $newFolder = $outputFolder + $_.Name # define folder as file name
      if( -not (Test-Path $newFolder) ) {  
        New-Item $newFolder -Type Directory # create with new Foldername
        if (! $newFolder.EndsWith('\')) { $newFolder += '\' } # and add \ to end      
        for($i = 1 ; $i -le $reader.NumberOfPages ; $i++) {
            # combine new filepath in output folder into directory name as filename
            $newFileName = $newFolder + $_.BaseName + ' -page- ' + $i + $_.Extension
            $document = New-Object iTextSharp.text.Document # instance of document
            # new instance of Stream with create method
            $fileStream = New-Object System.IO.FileStream($newFileName, [System.IO.FileMode]::Create)
            # new pdfCopy object with arguments of document and stream instances 
            $pdfCopy = New-Object iTextSharp.text.pdf.PdfCopy($document, $fileStream)
            $document.Open() # open document
            $pdfCopy.AddPage($pdfCopy.GetImportedPage($reader, $i)) # add page by $i

            $pdfCopy.Dispose() # dispose all instances
            $fileStream.Dispose()
            $document.Dispose() } <# end of FOR loop #> }  } # end of IF 
      if ($reader.NumberOfPages -eq 1 ) { 
        #Copy-Item $_.FullName -Destination $outputFolder 
        $newFileName = $outputFolder + $_.Name
        PDF-Merge-Process $inputFolder -fileName $newFileName -filter $_.Name } <# else just copy #> 
      $reader.Dispose() } # end of foreach
    # start output folder if all checks is OK
    Start-Process explorer -WindowStyle Maximized -ArgumentList $outputFolder }
} # End of Split-Pdf function

# WINDOW
[xml]$XAML = @"

<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:EP_PDF_TOOLS"
        WindowStartupLocation="CenterScreen"
        Title="EP PDF TOOLS v.0.1.0a" Height="225" MinHeight="225" MaxHeight="225" Width="400" MaxWidth="800" MinWidth="400" ResizeMode="CanResizeWithGrip">

    <Grid>
        <GroupBox Header="Select input folder" Height="70" VerticalAlignment="Top">
            <Grid Margin="0,0,0,0">
                <TextBox x:Name="input_folder_text" ScrollViewer.HorizontalScrollBarVisibility="Visible" TextWrapping="NoWrap" AcceptsReturn="True" Margin="5,3,50,3"/>
                <Button x:Name="browse_input_folder" Content="Ì" Margin="0,3,3,3" HorizontalAlignment="Right" Width="40" FontFamily="Webdings" FontSize="25" Foreground="#ccb000"/>
            </Grid>
        </GroupBox>
        <GroupBox Header="Select output folder" Margin="0,70,0,0" Height="70" VerticalAlignment="Top">
            <Grid Margin="0,0,0,0">
                <TextBox x:Name="output_folder_text" ScrollViewer.HorizontalScrollBarVisibility="Visible" TextWrapping="NoWrap" AcceptsReturn="True" Margin="5,3,50,3"/>
                <Button x:Name="browse_output_folder" Content="Ì" Margin="0,3,3,3" HorizontalAlignment="Right" Width="40" FontSize="25" FontFamily="Webdings" Foreground="#ccb000"/>
            </Grid>
        </GroupBox>
        <Button x:Name="split_button" Margin="0,0,15,10" Width="150" Height="30" VerticalAlignment="Bottom" HorizontalAlignment="Right" Content="SPLIT..." FontFamily="Segoe UI Symbol" FontSize="16"/>
        <Button x:Name="merge_button" Margin="10,0,15,10" Width="150" Height="30" VerticalAlignment="Bottom" HorizontalAlignment="Left" Content="MERGE..." FontFamily="Segoe UI Symbol" FontSize="16"/>
    </Grid>
</Window>

"@

# CONTROLS
$reader=(New-Object System.Xml.XmlNodeReader $XAML)
$Window=[Windows.Markup.XamlReader]::Load( $reader )
$XAML.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | 
ForEach { New-Variable  -Name $_.Name -Value $Window.FindName($_.Name) -Force }

# DEFAULT FOLDERS
$input_folder_text.Text = $HOME + '\Desktop\'
$output_folder_text.Text = $HOME + '\Downloads\'

# EVENTS
$browse_input_folder.Add_Click({ $input_folder_text.Text = Get-Folder })
$browse_output_folder.Add_Click({ $output_folder_text.Text = Get-Folder })
$split_button.Add_Click({ 
    Split-PDF $input_folder_text.Text $output_folder_text.Text
    $input_folder_text.Text = $output_folder_text.Text })
$merge_button.Add_Click({ Merge-PDF $input_folder_text.Text $output_folder_text.Text })

# STARTUP
$Null = $Window.ShowDialog()