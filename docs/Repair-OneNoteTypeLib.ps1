[CmdletBinding(ConfirmImpact='High',
  SupportsShouldProcess=$true,
  PositionalBinding=$false)]
Param
(
  [switch]$Force,
  [switch]$Undo
)

$tlb = [System.IO.Path]::Combine($PSScriptRoot, '..', 'idl', 'onenote15fix.tlb');
$profile = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile);
$description = $null;
$warning = $null;
$title = $null;

If ($Undo)
{
  $description = "Unregister onenote15fix.tlb for the current user and remove it from user profile.`n(DO NOT use if OneNote is installed on a per-user basis for the current user!)";
  $warning = "Do you want to register it for the current user and remove it from user profile?`n(DO NOT use if OneNote is installed on a per-user basis for the current user!)";
  $title = 'Undo repair OneNote type library for current user';
}
Else
{
  $description = "Copy onenote15fix.tlb to the user profile and register it for the current user.`n(DO NOT use if OneNote is installed on a per-user basis for the current user!)";
  $warning = "Do you want to copy onenote15fix.tlb to the user profile and register it for the current user?`n(DO NOT use if OneNote is installed on a per-user basis for the current user!)";
  $title = 'Repair OneNote type library for current user';
  If (-not (Test-Path -PathType 'Leaf' -LiteralPath $tlb))
  {
    Write-Error -Message 'Could not find onenote15fix.tlb.' `
      -RecommendedAction 'Use MIDL.EXE to compile onenote15fix.idl.' `
      -Category 'ObjectNotFound' -ErrorId 'RepairOneNoteTypeLib.NoTLB';
    Return;
  }
  If (-not (Test-Path -PathType 'Container' -LiteralPath $profile))
  {
    Write-Error -Message 'Profile folder is not found.' `
      -Category 'PermissionDenied' -ErrorId 'RepairOneNoteTypeLib.ProfileFolder';
    Return;
  }
}

If ($PSBoundParameters.ContainsKey('WhatIf') -and $PSBoundParameters['WhatIf'])
{
  $Force = $false;
}
If (-not $Force -and
  -not $PSCmdlet.ShouldProcess($description, $warning, $title))
{
  Return;
}

If ($Undo)
{
  $tlb = [System.IO.Path]::Combine($profile, 'onenote15fix.tlb');
  Remove-Item -LiteralPath 'HKCU:\SOFTWARE\Classes\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}' -Force -Recurse;
  Remove-Item -LiteralPath 'HKCU:\SOFTWARE\Classes\WOW6432Node\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}' -Force -Recurse;
  Remove-Item -LiteralPath $tlb -Force;
  Write-Verbose 'Repair-OneNoteTypeLib completed.';
  Return;
}

Copy-Item -LiteralPath $tlb -Destination $profile -Force | Out-Null;

$tlb = [System.IO.Path]::Combine($profile, 'onenote15fix.tlb');

New-Item -Path 'HKCU:\SOFTWARE\Classes\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}\1.1' -Value 'Microsoft OneNote 15.0 Type Library (fixed)' -Force | Out-Null;

New-Item -Path 'HKCU:\SOFTWARE\Classes\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}\1.1\0\Win32' -Value $tlb -Force | Out-Null;

New-Item -Path 'HKCU:\SOFTWARE\Classes\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}\1.1\Flags' -Value '0' -Force | Out-Null;

New-Item -Path 'HKCU:\SOFTWARE\Classes\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}\1.1\HelpDir' -Value '' -Force | Out-Null;

New-Item -Path 'HKCU:\SOFTWARE\Classes\WOW6432Node\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}\1.1' -Value 'Microsoft OneNote 15.0 Type Library (fixed)' -Force | Out-Null;

New-Item -Path 'HKCU:\SOFTWARE\Classes\WOW6432Node\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}\1.1\0\Win32' -Value $tlb -Force | Out-Null;

New-Item -Path 'HKCU:\SOFTWARE\Classes\WOW6432Node\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}\1.1\Flags' -Value '0' -Force | Out-Null;

New-Item -Path 'HKCU:\SOFTWARE\Classes\WOW6432Node\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}\1.1\HelpDir' -Value '' -Force | Out-Null;

Write-Verbose -Message 'Repair-OneNoteTypeLib completed.';
