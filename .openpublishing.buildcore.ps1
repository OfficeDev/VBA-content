param(
    [string]$parameters
)

# Main
$errorActionPreference = 'Stop'

# Entry-point package
$entryPointPackage = @{
    id = "opbuild.scripts";
    version = "latest";
    targetFramework = "net45";
}

# Pre-step: Set the repository root folder and working folder
$repositoryRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$workingDirectory = "$repositoryRoot\.optemp"

# Define system value
$systemDefaultVariables = @{
    ResourceContainerUrl = "https://opbuildstoragesandbox2.blob.core.windows.net/opps1container";
    DefaultEntryPoint = "op";
    UpdateNugetExe = $false;
    UpdateNugetConfig = $true;
    UpdateMdproj = $true;
    NeedBuildMdproj = $true;
    MdprojTargets = "build";
    OutputFolder = "$repositoryRoot\_site";
    LogOutputFolder = "$repositoryRoot\log";
    CacheFolder = "$workingDirectory\cache";
    LogLevel = "Info";
    NeedGeneratePdf = $false;
    UpdatePackagesConfig = $true;
    GlobalMetadataFile = "";
    DefaultMaxRetryCount = 3;
    NeedFetchSubmodule = $true;
    DefaultSubmoduleBranch = "master";
    DownloadNugetExeTimeOutInSeconds= 300;
    DownloadNugetConfigTimeOutInSeconds= 30;
    BuildToolParallelism = 0;
}
echo "default system value:" $systemDefaultVariables

Function ParseBoolValue([string]$variableName, [string]$stringValue, [bool]$defaultBoolValue)
{
    if([string]::IsNullOrEmpty($stringValue))
    {
        return $defaultBoolValue
    }

    try
    {
        $parsedBoolValue = [System.Convert]::ToBoolean($stringValue)
    }
    catch
    {
        Write-Error "variable $variableName does not have a valid bool value: $stringValue. Exception:$_.Exception.Message"
    }

    return $parsedBoolValue
}

Function GetValueFromVariableName([string]$variableValue, [string]$defaultStringValue)
{
    if([string]::IsNullOrEmpty($variableValue))
    {
        $variableValue = $defaultStringValue
    }
    return $variableValue
}

Function ParseParameters([string]$parameters)
{
    if([string]::IsNullOrEmpty($parameters))
    {
        return
    }

    $parameterPortions = $parameters.Split(';')
    foreach ($parameterPortion in $parameterPortions)
    {
        $keyValuePair = $parameterPortion.Split('=')
        if ($keyValuePair.Length -eq 2)
        {
            New-Variable -Name $keyValuePair[0] -Value $keyValuePair[1] -Scope "Global" -Force
            Write-Host "Create global variable with input $keyValuePair"
        }
        else
        {
            Write-Host "Invalid variable with input $keyValuePair. Ignore it."
        }
    }
}

Function IsPathExists([string]$path)
{
    return Test-Path $path
}

Function CheckPath([string]$path)
{
    if(!(IsPathExists($path)))
    {
        Write-Error "$path doesn't exist"
    }
}

Function CreateFolderIfNotExists([string]$folder)
{
    if(!(Test-Path "$folder"))
    {
        New-Item "$folder" -ItemType Directory
    }
}

Function RetryCommand
{
    param (
        [Parameter(Mandatory=$true)][string]$command,
        [Parameter(Mandatory=$true)][hashtable]$args,
        [Parameter(Mandatory=$false)][int]$maxRetryCount = $systemDefaultVariables.Item("DefaultMaxRetryCount"),
        [Parameter(Mandatory=$false)][ValidateScript({$_ -ge 0})][int]$retryIncrementalIntervalInSeconds = 10
    )

    # Setting ErrorAction to Stop is important. This ensures any errors that occur in the command are 
    # treated as terminating errors, and will be caught by the catch block.
    $args.ErrorAction = "Stop"

    $currentRetryIteration = 1
    $retryIntervalInSeconds = 0

    Write-Host ("Start to run command [{0}] with args [{1}]." -f $command, $($args | Out-String))
    do{
        try
        {
            Write-Host "Calling iteration $currentRetryIteration"
            & $command @args

            Write-Host "Command ['$command'] succeeded at iteration $currentRetryIteration."
            return
        }
        Catch
        {
            Write-Host "Calling iteration $currentRetryIteration failed, exception: '$($_.Exception.Message)'"
        }

        if ($currentRetryIteration -ne $maxRetryCount)
        {
            $retryIntervalInSeconds += $retryIncrementalIntervalInSeconds
            Write-Host "Command ['$command'] failed. Retrying in $retryIntervalInSeconds seconds."
            Start-Sleep -Seconds $retryIntervalInSeconds
        }
    } while (++$currentRetryIteration -le $maxRetryCount)

    Write-Host "Command ['$command'] failed. Maybe the network issues, please retry the build later."
    exit 1
}

Function DownloadFile([string]$source, [string]$destination, [bool]$forceDownload, [int]$timeoutSec = -1)
{
    if($forceDownload -or !(IsPathExists($destination)))
    {
        Write-Host "Download file to $destination from $source with force: $forceDownload"
        $destinationFolder = Split-Path -Parent $destination
        CreateFolderIfNotExists($destinationFolder)
        if ($timeoutSec -lt 0)
        {
            RetryCommand -Command 'Invoke-WebRequest' -Args @{ Uri = $source; OutFile = $destination; }
        }
        else
        {
            RetryCommand -Command 'Invoke-WebRequest' -Args @{ Uri = $source; OutFile = $destination; TimeoutSec = $timeoutSec }
        }
    }
}

Function GetPackageLatestVersion([string]$nugetExeDestination, [string]$packageName, [string]$nugetConfigDestination, [bool]$usePrereleasePackage = $false)
{
    $currentRetryIteration = 0;
    $maxRetryCount = $systemDefaultVariables.Item("DefaultMaxRetryCount");
    $retryIntervalInSeconds = 0;
    $retryIncrementalIntervalInSeconds = 10;

    do
    {
        Try
        {
            Write-Host "Use prerelease package for $packageName : $usePrereleasePackage"

            if ($usePrereleasePackage)
            {
                $filteredPackages = (& "$nugetExeDestination" list $packageName -Prerelease -ConfigFile "$nugetConfigDestination") -split "\n"
            }
            else
            {
                $filteredPackages = (& "$nugetExeDestination" list $packageName -ConfigFile "$nugetConfigDestination") -split "\n"
            }

            if ($LASTEXITCODE -eq 0)
            {
                foreach ($filteredPackage in $filteredPackages)
                {
                    $segments = $filteredPackage -split " "
                    if ($segments.Length -eq 2 -and $segments[0] -eq $packageName)
                    {
                        return $segments[1]
                    }
                }
            }

            Write-Host "Call iteration '$currentRetryIteration', cannot find latest version for $packageName, filtered packages: $filteredPackages"
        }
        Catch
        {
            Write-Host "Call iteration '$currentRetryIteration', cannot find latest version for $packageName, exception: $_.Exception.Message"
        }

        if ($currentRetryIteration -ne $maxRetryCount)
        {
            $retryIntervalInSeconds += $retryIncrementalIntervalInSeconds
            Write-Host "List package version failed, sleep $retryIntervalInSeconds seconds..."
            Start-Sleep -Seconds $retryIntervalInSeconds
        }
    } while (++$currentRetryIteration -le $maxRetryCount)

    Write-Host "Current nuget package list service is busy, please retry the build in 10 minutes"
    exit 1
}

Function RestorePackage([string] $nugetExeDestination, [string]$packagesDestination, [string]$packagesDirectory, [string]$nugetConfigDestination)
{
    Try
    {
        & "$nugetExeDestination" restore "$packagesDestination" -PackagesDirectory "$packagesDirectory" -ConfigFile "$nugetConfigDestination"
        return $LASTEXITCODE -eq 0
    }
    Catch
    {
        return $false;
    }
}

Function GeneratePackagesConfig([string]$outputFilePath, [object[]]$dependencies)
{
    $packageConfigXmlTemplate = @'
<?xml version="1.0" encoding="utf-8"?>
<packages></packages>
'@

    $packageConfigXml = [xml]$packageConfigXmlTemplate
    foreach ($dependency in $dependencies)
    {
        $packageNode = $packageConfigXml.CreateElement("package")
        $packageNode.SetAttribute("id", $dependency.id)
        
        if ($dependency.version -eq "latest" -or $dependency.version -eq "latest-prerelease")
        {
            $usePrereleasePackage = $dependency.version -eq "latest-prerelease"

            # Get latest package version
            $dependency.actualVersion = GetPackageLatestVersion($nugetExeDestination) ($dependency.id) ($nugetConfigDestination) ($usePrereleasePackage)

            Write-Host "Using version $($dependency.actualVersion) for package $($dependency.id) (requested: $($dependency.version))"
        }
        else
        {
            $dependency.actualVersion = $dependency.version
        }
        $packageNode.SetAttribute("version", $dependency.actualVersion)

        $packageNode.SetAttribute("targetFramework", $dependency.targetFramework)
        $packageConfigXml.SelectSingleNode("packages").AppendChild($packageNode)
    }
    
    if (IsPathExists($outputFilePath))
    {
        Remove-Item $outputFilePath -Force
    }
    $packageConfigXml.Save($outputFilePath)
}

Function ConvertToJsonSafely {
    param([string]$content)
    process { $_ | ConvertTo-Json -Depth 99 }
}

Filter timestamp
{
    if (![string]::IsNullOrEmpty($_) -and ![string]::IsNullOrWhiteSpace($_))
    {
        Write-Host -NoNewline -ForegroundColor Magenta "[$(((get-date).ToUniversalTime()).ToString("HH:mm:ss.ffffffZ"))]: "
    }

    $_
}

# Step-1: Parse parameters
echo "Parse parameters $parameters" | timestamp
ParseParameters($parameters)

# Step-2: Parse publish configuration
$publishConfigFile = "$repositoryRoot\.openpublishing.publish.config.json"
CheckPath($publishConfigFile)
$publishConfigContent = (Get-Content $publishConfigFile -Raw) | ConvertFrom-Json

# Step-3: Download Nuget tools and nuget config
echo "Download Nuget tool and config" | timestamp
$resourceContainerUrl = GetValueFromVariableName($resourceContainerUrl) ($systemDefaultVariables.Item("ResourceContainerUrl"))
$nugetConfigSource = "$resourceContainerUrl/Tools/Nuget/Nuget.Config"
$nugetExeSource = "$resourceContainerUrl/Tools/Nuget/nuget.exe"

$nugetConfigDestination = "$workingDirectory\Tools\Nuget\Nuget.Config"
$nugetExeDestination = "$workingDirectory\Tools\Nuget\nuget.exe"

$DownloadNugetExeTimeOutInSeconds = GetValueFromVariableName($DownloadNugetExeTimeOutInSeconds) ($systemDefaultVariables.Item("DownloadNugetExeTimeOutInSeconds"))
$DownloadNugetConfigTimeOutInSeconds = GetValueFromVariableName($DownloadNugetConfigTimeOutInSeconds) ($systemDefaultVariables.Item("DownloadNugetConfigTimeOutInSeconds"))
$UpdateNugetExe = ParseBoolValue("UpdateNugetExe") ($UpdateNugetExe) ($systemDefaultVariables.Item("UpdateNugetExe"))
DownloadFile($nugetExeSource) ($nugetExeDestination) ($UpdateNugetExe) ($DownloadNugetExeTimeOutInSeconds)
$UpdateNugetConfig = ParseBoolValue("UpdateNugetConfig") ($UpdateNugetConfig) ($systemDefaultVariables.Item("UpdateNugetConfig"))
DownloadFile($nugetConfigSource) ($nugetConfigDestination) ($UpdateNugetConfig) ($DownloadNugetConfigTimeOutInSeconds)

# Step-4: Create packages.config for entry-point package
echo "Create packages.config for entry-point package" | timestamp
$configPackageVersion = $publishConfigContent.package_version
if (![string]::IsNullOrEmpty($configPackageVersion))
{
    $entryPointPackage.version = $configPackageVersion
}

# for non-PROD env, treat latest version as latest-prerelease version by default
$treatLatestVersionAsLatestPrereleaseVersion = !$resourceContainerUrl.StartsWith("https://opbuildstorageprod.blob.core.windows.net")
if ($_op_treatLatestVersionAsLatestPrereleaseVersion)
{
    $treatLatestVersionAsLatestPrereleaseVersion = $_op_treatLatestVersionAsLatestPrereleaseVersion -eq "true"
}

if ($treatLatestVersionAsLatestPrereleaseVersion -and $entryPointPackage.version -eq "latest")
{
    $entryPointPackage.version = "latest-prerelease"
    echo "Use latest-prerelease version instead of latest version." | timestamp
}

$packagesDestination = "$workingDirectory\packages.config"
GeneratePackagesConfig($packagesDestination) (@($entryPointPackage))

# Step-5 Restore entry-point package
echo "Restore entry-point package: $($entryPointPackage.id)" | timestamp
$packagesDirectory = "$workingDirectory\packages"
$restoreSucceeded = RestorePackage($nugetExeDestination) ($packagesDestination) ($packagesDirectory) ($nugetConfigDestination)
if (!$restoreSucceeded)
{
    echo "Restore entry-point package failed" | timestamp
    exit 1
}

# Step-6: Call build entry point
$packageToolsDirectory = "$packagesDirectory\$($entryPointPackage.id).$($entryPointPackage.actualVersion)\tools"
$buildEntryPointDestination = "$packageToolsDirectory\build.entrypoint.ps1"
echo "Call build entry point at $buildEntryPointDestination" | timestamp
& "$buildEntryPointDestination" "$repositoryRoot" "$packagesDirectory"

exit $LASTEXITCODE
