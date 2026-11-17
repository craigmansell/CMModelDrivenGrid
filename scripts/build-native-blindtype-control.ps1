[CmdletBinding()]
param([string]$BuildRoot='')
$ErrorActionPreference='Stop'
$source=Split-Path -Parent $PSScriptRoot
if(!$BuildRoot){$BuildRoot=Join-Path (Split-Path -Parent $source) 'native-blindtype-build-20261117'}
$BuildRoot=[IO.Path]::GetFullPath($BuildRoot)
$workspace='C:\Users\CraigMansell\Documents\wubl\CrmPackage'
if(!$BuildRoot.StartsWith($workspace+'\',[StringComparison]::OrdinalIgnoreCase) -or $BuildRoot.StartsWith($source+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Use a separate native build directory inside the authorized workspace'}
$control=Join-Path $BuildRoot 'control';$solution=Join-Path $BuildRoot 'solution'
foreach($path in @($control,(Join-Path $control 'strings'),(Join-Path $solution 'src/Other'))){New-Item -ItemType Directory -Path $path -Force|Out-Null}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'native-blindtype-control/ControlManifest.template.xml') -Destination (Join-Path $control 'ControlManifest.Input.xml')
$relative=[IO.Path]::GetRelativePath($control,(Join-Path $PSScriptRoot 'native-blindtype-control/index')).Replace('\','/')
[IO.File]::WriteAllText((Join-Path $control 'index.ts'),"export {CMBlindTypeSetupGridV2} from '$relative';",[Text.UTF8Encoding]::new($false))
[xml]$resx=Get-Content (Join-Path $source 'CMModelDrivenGrid/strings/CMModelDrivenGrid.1033.resx') -Raw
$data=$resx.CreateElement('data');$data.SetAttribute('name','Blind_Type_Setup');$value=$resx.CreateElement('value');$value.InnerText='Blind type setup';[void]$data.AppendChild($value);[void]$resx.root.AppendChild($data);$resx.Save((Join-Path $control 'strings/CMBlindTypeSetupGrid.1033.resx'))
$project=[IO.File]::ReadAllText((Join-Path $source 'CMModelDrivenGrid.pcfproj')).Replace('<Name>CMModelDrivenGrid</Name>','<Name>CMBlindTypeSetupGridV2</Name>').Replace('Version="1.*"','Version="1.52.1"')
[IO.File]::WriteAllText((Join-Path $BuildRoot 'CMBlindTypeSetupGrid.pcfproj'),$project,[Text.UTF8Encoding]::new($false))
$package=Get-Content (Join-Path $source 'package.json') -Raw|ConvertFrom-Json;$package.name='blind-type-setup-grid';$package|ConvertTo-Json -Depth 20|Set-Content (Join-Path $BuildRoot 'package.json') -Encoding utf8
Copy-Item -LiteralPath (Join-Path $source 'package-lock.json') -Destination (Join-Path $BuildRoot 'package-lock.json')
if(!(Test-Path -LiteralPath (Join-Path $BuildRoot 'node_modules'))){New-Item -ItemType Junction -Path (Join-Path $BuildRoot 'node_modules') -Target (Join-Path $source 'node_modules')|Out-Null}
@{extends=((Join-Path $source 'node_modules/pcf-scripts/tsconfig_base.json').Replace('\','/'));compilerOptions=@{typeRoots=@((Join-Path $source 'node_modules/@types').Replace('\','/'))};include=@('control/**/*.ts')}|ConvertTo-Json -Depth 10|Set-Content (Join-Path $BuildRoot 'tsconfig.json') -Encoding utf8
@{outDir='./out/controls'}|ConvertTo-Json|Set-Content (Join-Path $BuildRoot 'pcfconfig.json') -Encoding utf8
$lint=Join-Path $source 'eslint.config.mjs';if(Test-Path $lint){[IO.File]::WriteAllText((Join-Path $BuildRoot 'eslint.config.mjs'),"export {default} from '$($lint.Replace('\','/'))';",[Text.UTF8Encoding]::new($false))}
$solutionProject=[IO.File]::ReadAllText((Join-Path $source 'CMModelDrivenGridSolution/CMModelDrivenGridSolution.cdsproj')).Replace('..\CMModelDrivenGrid.pcfproj','..\CMBlindTypeSetupGrid.pcfproj').Replace('Version="1.*"','Version="1.52.1"')
[IO.File]::WriteAllText((Join-Path $solution 'CMBlindTypeSetup.cdsproj'),$solutionProject,[Text.UTF8Encoding]::new($false))
foreach($name in @('Customizations.xml','Relationships.xml')){Copy-Item -LiteralPath (Join-Path $source "CMModelDrivenGridSolution/src/Other/$name") -Destination (Join-Path $solution "src/Other/$name")}
[xml]$manifest=Get-Content (Join-Path $source 'CMModelDrivenGridSolution/src/Other/Solution.xml') -Raw
$manifest.ImportExportXml.SolutionManifest.UniqueName='CMBlindTypeSetup';$manifest.ImportExportXml.SolutionManifest.Version='1.1.0.0';$manifest.ImportExportXml.SolutionManifest.LocalizedNames.LocalizedName.SetAttribute('description','Blind Type Setup');$manifest.Save((Join-Path $solution 'src/Other/Solution.xml'))
Push-Location $BuildRoot
try{dotnet build (Join-Path $solution 'CMBlindTypeSetup.cdsproj') --configuration Release /p:PcfEnableAutoNpmInstall=false /p:SolutionPackageType=Unmanaged --ignore-failed-sources;if($LASTEXITCODE -ne 0){throw 'Native PCF build/package failed'}}finally{Pop-Location}
Write-Output (Join-Path $solution 'bin/Release/CMBlindTypeSetup.zip')
