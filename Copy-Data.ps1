function Start-Copy {

    # Use the Get-Folder function to select the source and destination folder through a GUI dialog box
    $Source = Get-Folder("Select source folder");
    $Destination = Get-Folder ("Select destination folder");

    Write-Verbose -Message ('Source: {0}' -f $Source);
    Write-Verbose -Message ('Destination: {0}' -f $Destination);

    # When no source or destination is selected inform the user
    if ($Source -eq "" -OR $Destination -eq "") {
        Write-Host "No source or destination selected."
    } else {
        Copy-WithProgress -Source $Source -Destination $Destination
    }
}

function Get-Folder ($DialogMessage) {
    <#
    .NAME
        Get-Folder
    .SYNOPSIS
        Returns a folder path.
    .DESCRIPTION
        Shows a GUI dialog box to select then return a folder path.
    .PARAMETER DialogMessage
        Provide a message to display in the dialog box.
    #>

    # Import windows forms for a GUI folder selection
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

    # Store the folder path for returning
    $FolderPath = ""

    # Display a dialog box to select a source folder
    $Dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $Dialog.Description = $DialogMessage
    $Dialog.RootFolder = "MyComputer"

    # Only write to FolderPath if the OK button is clicked in the dialog window
    if ($Dialog.ShowDialog() -eq "OK") {
        $FolderPath = $Dialog.SelectedPath
    }

    return $FolderPath
}


function Copy-WithProgress {
    
    <#
    .NAME
        Copy-WithProgress
    .SYNOPSIS
        Copy a file or directory from a Source to a Destination.
    .DESCRIPTION
        Copy a file of direction using the built-in robocopy command from a specified Source to a Destination.
    .PARAMETER Source
        The source to be copied.
    .PARAMETER Destination
        The destionation to copy to.
    .PARAMETER Delay
        The delay between file copies to allow calculations.
    .PARAMETER ReportDelay
        The delay between each report to the progress bar.
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Source
       ,[Parameter(Mandatory = $true)]
        [string] $Destination
       ,[int] $Delay = 200
       ,[int] $ReportDelay = 2000
    )

    # Regular expression to gather number of bytes copied
    $RegexBytes = '(?<=\s+)\d+(?=\s+)';

    # Retrieve the source directory name for use in the destination path
    $SourceFolder = Split-Path -Path $Source -Leaf;
    $DestinationFolder = $Destination;
    $Destination = ('{0}\{1}' -f $Destination, $SourceFolder);

    #region Robocopy parameters
    # /e - Copies subdirectories. This option automatically includes empty directories.
    # /xj - Excludes junction points, which are normally included by default.
    # /np - Specifies that the progress of the copying operation (the number of files or directories copied so far) will not be displayed.
    # /nc - Specifies that file classes are not to be logged.
    # /ndl - Specifies that directory names are not to be logged.
    # /bytes - Prints sizes, as bytes.
    # /njh - Specifies that there is no job header.
    # /njs - Specifies that there is no job summary.
    # /tee - Writes the status output to the console window, as well as to the log file.
    $RobocopyParams = '/e /xj /np /nc /ndl /bytes /njh /njs /r:1 /w:0';
    #endregion Robocopy params

    #region Robocopy analysis
    # Use the built-in RoboCopy command to parse a provided Source to determine the total number of files as will as the total size of the Source using a temporary log
    Write-Host 'Calculating robocopy job...';
    $AnalysisLogPath = '{0}\temp\{1} RobocopyAnalysis.log' -f $env:windir, (Get-Date -Format 'yyyy-MM-dd HH-mm-ss');

    # Parse source with robocopy command
    $AnalysisArgumentList = '"{0}" "{1}" /log:"{2}" /l {3}' -f $Source, $Destination, $AnalysisLogPath, $RobocopyParams;
    Start-Process -Wait -FilePath RoboCopy.exe -ArgumentList $AnalysisArgumentList -NoNewWindow;

    # Display provided Source and Destination
    Write-Host
    Write-Host ('Source: {0}' -f $Source);
    Write-Host ('Destination: {0}' -f $Destination);

    # Get the total number of files that will be copied
    $AnalysisContent = Get-Content -Path $AnalysisLogPath;
    $TotalFileCount = $AnalysisContent.Count - 1;
    Write-Host
    Write-Host ('Total number of files that will be copied: {0}' -f $TotalFileCount);

    # Get the total number of bytes that will be copied
    [RegEx]::Matches(($AnalysisContent -join "`n"), $RegexBytes) | % { $BytesTotal = 0; } { $BytesTotal += $_.Value; };
    $GigabytesTotal = [math]::Round($BytesTotal/1073741824,3);
    Write-Host ('Total number of Gigabytes that will be copied: {0} GB' -f $GigabytesTotal);
    #endregion Robocopy analysis

    #region Robocopy transfer
    # Use the built-in Robocopy command to transfer all files from the provided Source to the provided Destination

    # Create a log file in the Windows Temp directory
    $RobocopyLogPath = '{0}\CopyLog-{1} {2}.log' -f $DestinationFolder, $SourceFolder , (Get-Date -Format 'yyyy-MM-dd HH-mm-ss');
    # Configure argument list for use when running the RoboCopy command
    $RobocopyArgumentList = '"{0}" "{1}" /log:"{2}" /ipg:{3} {4}' -f $Source, $Destination, $RobocopyLogPath, $Delay, $RobocopyParams;

    # Begin RoboCopy transfer from the provided Source to the provided Destination
    Write-Host
    Write-Verbose -Message ('Beginning the robocopy process with the arguments: {0}' -f $RobocopyArgumentList);
    $RobocopyTransfer = Start-Process -FilePath RoboCopy.exe -ArgumentList $RobocopyArgumentList -PassThru -NoNewWindow;
    Start-Sleep -Milliseconds 100;
    #endregion Robocopy transfer

    #region Progress bar
    # Display a progress bar in the powershell window to visually show how far along the RoboCopy transfer is

    # Loop the progress bar until the RoboCopy transfer is complete
    while (!$RobocopyTransfer.HasExited) {
        Start-Sleep -Milliseconds $ReportDelay
        $BytesCopied = 0;
        $TransferLogContent = Get-Content -Path $RobocopyLogPath;
        $BytesCopied = [Regex]::Matches($TransferLogContent, $RegexBytes) | ForEach-Object -Process { $BytesCopied += $_.Value; } -End { $BytesCopied; };
        $CopiedFileCount = $TransferLogContent.Count - 1;

        # Write how much data and how many files have been copied so far
        Write-Verbose -Message ('Bytes copied: {0}' -f $BytesCopied);
        Write-Verbose -Message ('Files copied: {0}' -f $TransferLogContent.Count);

        # Calculate the percentage transferred so far
        $Percentage = 0;
        if ($BytesCopied -gt 0) {
            $Percentage = (($BytesCopied/$BytesTotal)*100)
        }

        # Convert bytes to gigabytes
        $GigabytesCopied = [math]::Round(($BytesCopied/1073741824),3);

        # Display progress bar
        Write-Progress -Activity RobocopyTransfer -Status ("Copied: {0} of {1} files | Copied: {2} of {3} GB | Percent complete: {4}%" -f $CopiedFileCount, $TotalFileCount, $GigabytesCopied, $GigabytesTotal, [math]::Round($Percentage,2)) -PercentComplete $Percentage
    }
    
    # Display progress bar as complete
    Write-Progress -Activity RobocopyTransfer -Status "Ready" -Completed
    #endregion Progress bar

    #region Function summary
    # Clear currently displayed progress bar and display a summary of the RoboCopy transfer

    Clear-Host

    # Only display summary if files are actually transferred
    if ($Percentage -gt 0.1) {
        Write-Host "Done transfering all files!"
        Write-Host
        Write-Host ('Source: {0}' -f $Source);
        Write-Host ('Destination: {0}' -f $Destination);
        Write-Host
        Write-Host ('Number of files: {0} of {1}' -f $CopiedFileCount, $TotalFileCount);
        Write-Host ('Gigabytes of data: {0} of {1}' -f $GigabytesCopied, $GigabytesTotal);

    # If files are not transferred refer user to the logs
    } else {
        Write-Host "All files are most likely already at the destination, please check the log for details"
    }

    Write-Host
    Write-Host ('Check the log at {0} for further details' -f $RobocopyLogPath);
    #endregion Function summary

}