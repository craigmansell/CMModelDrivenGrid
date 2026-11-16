[CmdletBinding()]
param(
    [ValidateSet('Prepare','Apply','Publish','Verify','Export')][string]$Mode='Prepare',
    [string]$Baseline='',
    [string]$Package='',
    [string]$ValidatedPackageHash='',
    [string]$CrmWorkspace='C:\Users\CraigMansell\Documents\wubl\CrmPackage',
    [switch]$AllowPublishAll
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
. (Join-Path $CrmWorkspace "Migration/DataverseAuth.ps1")
$target='https://capricornsandbox.crm4.dynamics.com'
$org='27252def-2a75-f111-b27b-000d3ab8db73'
$controlId='7dc1b9c0-a252-4306-a78b-fb5020588250'
$controlName='cm_CraigMansell.CMModelDrivenGrid'
$solutionName='CMModelDrivenGridSolution'
$publisherId='acc5f823-21ea-f011-8406-002248a34d35'
$beforeVersion='1.0.117';$beforeSolutionVersion='1.0.96.0'
$version='1.0.119';$solutionVersion='1.0.98.0'
$fromVersion=$beforeVersion;$fromSolutionVersion=$beforeSolutionVersion
$artifactPrefix='v119'
$fromBundleHash='A7190A4DE5D9B41D66745372F518209135C3F54F62DFE4FB871F1614DB9B3151'
$hierarchyViewId='c2c8d376-3232-4b14-bb7a-0594df431861'
$token=Get-DataverseAccessToken -EnvUrl $target -Label 'DrivenGrid native presentation sandbox'
$api="$target/api/data/v9.2"
$headers=@{Authorization="Bearer $token";Accept='application/json';'OData-Version'='4.0'}
function Read-Dv([string]$Path){Invoke-RestMethod -Uri "$api/$Path" -Headers $headers}
function Post-Dv([string]$Path,$Body){Invoke-RestMethod -Uri "$api/$Path" -Headers $headers -Method Post -ContentType 'application/json' -Body ($Body|ConvertTo-Json -Depth 50 -Compress)}
function Save-Json([string]$Path,$Value){$Value|ConvertTo-Json -Depth 60|Set-Content -LiteralPath $Path -Encoding utf8}
function Hash-Text([string]$Value){[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($Value)))}
function Get-TextField($Object,[string]$Name){if($Object -is [System.Collections.IDictionary]){return [string]$Object[$Name]};$property=$Object.PSObject.Properties[$Name];if($null -eq $property){return ''};return [string]$property.Value}
function Read-Package([string]$Path){
    $zip=[IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $Path).Path)
    try{
        $reader=[IO.StreamReader]::new($zip.GetEntry('solution.xml').Open())
        try{[xml]$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
        $manifest=$xml.ImportExportXml.SolutionManifest
        if([string]$manifest.Managed -cne '0' -or [string]$manifest.UniqueName -cne $solutionName -or [string]$manifest.Publisher.UniqueName -cne 'CraigMansell'){throw 'Package unmanaged/identity/publisher lock failed'}
        $entries=@($zip.Entries|Where-Object {$_.FullName -match '(^|/)ControlManifest.xml$'})
        if($entries.Count -ne 1){throw 'Expected exactly one PCF manifest'}
        $reader=[IO.StreamReader]::new($entries[0].Open())
        try{[xml]$pcf=$reader.ReadToEnd()}finally{$reader.Dispose()}
        if($pcf.manifest.control.namespace -cne 'CraigMansell' -or $pcf.manifest.control.constructor -cne 'CMModelDrivenGrid'){throw 'PCF identity mismatch'}
        $assets=@{}
        foreach($asset in @('bundle.js','css/CMModelDrivenGrid.css','strings/CMModelDrivenGrid.1033.resx')){
            $entry=$zip.GetEntry("Controls/cm_CraigMansell.CMModelDrivenGrid/$asset")
            if(!$entry){throw "Missing candidate resource: $asset"}
            $stream=$entry.Open()
            try{$assets["cc_CraigMansell.CMModelDrivenGrid/$asset"]=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($stream))}finally{$stream.Dispose()}
        }
        return @{solutionVersion=[string]$manifest.Version;controlVersion=[string]$pcf.manifest.control.version;assets=$assets}
    }finally{$zip.Dispose()}
}
function Get-Forms([switch]$Draft){
    $path=if($Draft){'systemforms/Microsoft.Dynamics.CRM.RetrieveUnpublishedMultiple()'}else{'systemforms'}
    $result=Read-Dv ($path+'?$select=formid,name,formxml&$filter=objecttypecode eq ''cm_itemtemplate'' and type eq 2')
    if($result.PSObject.Properties.Name -contains '@odata.nextLink'){throw 'Form audit requires pagination'}
    return @($result.value)
}
function Assert-Forms($Before,$Now){
    if(@($Before).Count -ne @($Now).Count){throw 'Form count changed'}
    foreach($form in $Before){$found=@($Now|Where-Object formid -eq $form.formid);if($found.Count -ne 1 -or (Hash-Text $form.formxml) -cne (Hash-Text $found[0].formxml)){throw "Concurrent form change: $($form.name)"}}
}
function Get-Views([switch]$Draft){
    $path=if($Draft){'savedqueries/Microsoft.Dynamics.CRM.RetrieveUnpublishedMultiple()'}else{'savedqueries'}
    $result=Read-Dv ($path+'?$select=savedqueryid,name,fetchxml,layoutxml,querytype&$filter=returnedtypecode eq ''cm_itemtemplateoption''')
    if($result.PSObject.Properties.Name -contains '@odata.nextLink'){throw 'View audit requires pagination'}
    return @($result.value)
}
function Assert-Views($Before,$Now){
    if(@($Before).Count -ne @($Now).Count){throw 'View inventory changed'}
    foreach($view in $Before){
        $found=@($Now|Where-Object savedqueryid -eq $view.savedqueryid)
        if($found.Count -ne 1){throw 'View missing'}
        foreach($field in @('name','fetchxml','layoutxml','querytype')){
            if((Hash-Text (Get-TextField $view $field)) -cne (Hash-Text (Get-TextField $found[0] $field))){throw "Concurrent view change: $($view.name)"}
        }
    }
}
function Get-Pending{
    $pending=@{}
    foreach($entry in @(@('systemforms','formid,name'),@('webresourceset','webresourceid,name'),@('savedqueries','savedqueryid,name'),@('appmodules','appmoduleid,name'),@('savedqueryvisualizations','savedqueryvisualizationid,name'),@('sitemaps','sitemapid,sitemapname'))){
        $result=Read-Dv ($entry[0]+'/Microsoft.Dynamics.CRM.RetrieveUnpublishedMultiple()?$select='+$entry[1]+',componentstate&$filter=componentstate eq 1')
        if($result.PSObject.Properties.Name -contains '@odata.nextLink'){throw 'Pending audit requires pagination'}
        $pending[$entry[0]]=@($result.value)
    }
    return $pending
}
function Assert-Pending($Pending,[switch]$AllowOwn){
    foreach($key in $Pending.Keys){foreach($record in $Pending[$key]){
        if(!$AllowOwn -or $key -ne 'webresourceset' -or $record.name -notmatch '^cc_CraigMansell\.CMModelDrivenGrid/(bundle\.js|css/CMModelDrivenGrid\.css|strings/CMModelDrivenGrid\.1033\.resx)$'){throw "Pending customization in $key; stop for review"}
    }}
}
function Get-Solution{
    $rows=@((Read-Dv 'solutions?$select=solutionid,uniquename,version,ismanaged,_publisherid_value&$filter=uniquename eq ''CMModelDrivenGridSolution''').value)
    if($rows.Count -ne 1 -or $rows[0].ismanaged -or $rows[0]._publisherid_value -ne $publisherId){throw 'Installed solution unmanaged/publisher lock failed'}
    return $rows[0]
}
function Get-Publisher{Read-Dv "publishers($publisherId)?`$select=uniquename,customizationprefix,customizationoptionvalueprefix"}
function Assert-Publisher($Value){if($Value.uniquename -cne 'CraigMansell' -or $Value.customizationprefix -cne 'cm' -or $Value.customizationoptionvalueprefix -ne 17218){throw 'Publisher settings changed'}}
function Get-Control{Read-Dv "customcontrols($controlId)?`$select=name,version,manifest,clientjson"}
function Export-Pcf([string]$Path){
    $result=Post-Dv 'ExportSolution' @{SolutionName=$solutionName;Managed=$false}
    [IO.File]::WriteAllBytes($Path,[Convert]::FromBase64String($result.ExportSolutionFile))
    $manifest=Read-Package $Path
    return @{path=$Path;sha256=(Get-FileHash -LiteralPath $Path).Hash;solutionVersion=$manifest.solutionVersion;controlVersion=$manifest.controlVersion}
}
if((Read-Dv 'WhoAmI').OrganizationId -ne $org){throw 'Sandbox organization lock failed'}
$installed=Get-Control
if($installed.name -cne $controlName){throw 'Control identity mismatch'}
$solution=Get-Solution
Assert-Publisher (Get-Publisher)
$forms=Get-Forms
$views=Get-Views
[xml]$form=@($forms|Where-Object formid -eq '52d64db5-87fb-f011-8406-002248a34d35')[0].formxml
$subgrid=$form.SelectSingleNode('//control[@id="Subgrid_ItemTemplateOptions"]')
if(!$subgrid -or $subgrid.parameters.TargetEntityType -cne 'cm_itemtemplateoption' -or $subgrid.parameters.RelationshipName -cne 'cm_itemtemplateoption_ItemTemplate_cm_itemtemplate'){throw 'Subgrid identity/scope mismatch'}
if($subgrid.SelectSingleNode('parameters/ViewId').InnerText.Trim('{}') -ine $hierarchyViewId){throw 'Subgrid hierarchy view changed'}
$bindings=@($form.SelectNodes("//controlDescription[@forControl='$($subgrid.uniqueid)']/customControl[@name='$controlName']"))
if($bindings.Count -ne 3){throw 'Expected installed DrivenGrid binding on all three form factors'}
foreach($binding in $bindings){
    if($binding.parameters.EnableNativeChildSelection.InnerText -cne 'true' -or $binding.parameters.EnableChildRows.InnerText -cne 'true' -or $binding.parameters.ChildParentLookup.InnerText -cne 'cm_parenttemplateoption'){throw 'Native hierarchy settings changed'}
    if($binding.SelectSingleNode('parameters/data-set[@name="records"]/ViewId').InnerText.Trim('{}') -ine $hierarchyViewId){throw 'Native hierarchy dataset view changed'}
}
if($Mode -eq 'Prepare'){
    if($installed.version -cne $beforeVersion -or $solution.version -cne $beforeSolutionVersion){throw 'Installed baseline changed'}
    $pending=Get-Pending;Assert-Pending $pending
    Assert-Forms $forms (Get-Forms -Draft)
    Assert-Views $views (Get-Views -Draft)
    $directory=Join-Path $CrmWorkspace ('outputs/deployment-backups/drivengrid-native-look-sandbox-'+(Get-Date -Format yyyyMMdd-HHmmss))
    New-Item -ItemType Directory -Path $directory|Out-Null
    Save-Json (Join-Path $directory 'forms.before.json') $forms
    Save-Json (Join-Path $directory 'views.before.json') $views
    Save-Json (Join-Path $directory 'control.before.json') $installed
    Save-Json (Join-Path $directory 'publisher.before.json') (Get-Publisher)
    Save-Json (Join-Path $directory 'pending.before.json') $pending
    $resources=@((Read-Dv 'webresourceset?$select=webresourceid,name,content&$filter=startswith(name,''cc_CraigMansell.CMModelDrivenGrid/'')').value)
    if($resources.Count -ne 3){throw 'Unexpected PCF resource count'}
    Save-Json (Join-Path $directory 'resources.before.json') $resources
    $rollbackPackage=Export-Pcf (Join-Path $directory 'CMModelDrivenGridSolution.before.unmanaged.zip')
    if($rollbackPackage.controlVersion -cne $beforeVersion -or $rollbackPackage.solutionVersion -cne $beforeSolutionVersion){throw 'Rollback package differs from published baseline'}
    $rollbackAssets=(Read-Package $rollbackPackage.path).assets
    foreach($resource in $resources){
        $resourceHash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Convert]::FromBase64String($resource.content)))
        if(!$rollbackAssets.ContainsKey($resource.name) -or $rollbackAssets[$resource.name] -cne $resourceHash){throw 'Rollback resources differ from the published control'}
    }
    if($rollbackAssets['cc_CraigMansell.CMModelDrivenGrid/bundle.js'] -cne $fromBundleHash){throw 'Published baseline differs from the verified 1.0.117 bundle'}
    # Preserve the complete current release baseline before this PCF-only batch.
    $releasePath=Join-Path $directory 'CapricornRelease20260622.before.unmanaged.zip'
    $release=Post-Dv 'ExportSolution' @{SolutionName='CapricornRelease20260622';Managed=$false}
    [IO.File]::WriteAllBytes($releasePath,[Convert]::FromBase64String($release.ExportSolutionFile))
    $releaseZip=[IO.Compression.ZipFile]::OpenRead($releasePath)
    try {
        $releaseReader=[IO.StreamReader]::new($releaseZip.GetEntry('solution.xml').Open())
        try {[xml]$releaseXml=$releaseReader.ReadToEnd()}finally{$releaseReader.Dispose()}
        if($releaseXml.ImportExportXml.SolutionManifest.UniqueName -cne 'CapricornRelease20260622' -or $releaseXml.ImportExportXml.SolutionManifest.Managed -cne '0'){throw 'Full rollback package unmanaged/identity lock failed'}
    }finally{$releaseZip.Dispose()}
    $releaseRollback=@{path=$releasePath;sha256=(Get-FileHash -LiteralPath $releasePath).Hash;uniqueName='CapricornRelease20260622';managed=$false}
    Assert-Forms $forms (Get-Forms);Assert-Forms $forms (Get-Forms -Draft)
    Assert-Views $views (Get-Views);Assert-Views $views (Get-Views -Draft)
    Assert-Pending (Get-Pending)
    Save-Json (Join-Path $directory 'rollback-manifest.json') @{createdUtc=[DateTime]::UtcNow.ToString('o');target=$target;organizationId=$org;directory=$directory;package=$rollbackPackage;releasePackage=$releaseRollback;formHash=(Get-FileHash -LiteralPath (Join-Path $directory 'forms.before.json')).Hash;viewsHash=(Get-FileHash -LiteralPath (Join-Path $directory 'views.before.json')).Hash;controlVersion=$installed.version;solutionVersion=$solution.version}
    Write-Output (Join-Path $directory 'rollback-manifest.json')
    exit
}
if(!$Baseline){throw 'Fresh rollback baseline required'}
$saved=Get-Content -LiteralPath $Baseline -Raw|ConvertFrom-Json
$baselineCreated=[DateTimeOffset]$saved.createdUtc
if(([DateTime]::UtcNow-$baselineCreated.UtcDateTime).TotalHours -gt 12){throw 'Rollback baseline is older than this deployment batch'}
if($saved.target -cne $target -or $saved.organizationId -ne $org){throw 'Baseline target mismatch'}
if($saved.controlVersion -cne $beforeVersion -or $saved.solutionVersion -cne $beforeSolutionVersion){throw 'Expected rollback baseline versions differ'}
if((Get-FileHash -LiteralPath $saved.package.path).Hash -cne $saved.package.sha256){throw 'Rollback package changed'}
if((Get-FileHash -LiteralPath $saved.releasePackage.path).Hash -cne $saved.releasePackage.sha256){throw 'Full release rollback package changed'}
$rollbackIdentity=Read-Package $saved.package.path
if($rollbackIdentity.controlVersion -cne $beforeVersion -or $rollbackIdentity.solutionVersion -cne $beforeSolutionVersion){throw 'Rollback package versions differ'}
$formPath=Join-Path $saved.directory 'forms.before.json'
if((Get-FileHash -LiteralPath $formPath).Hash -cne $saved.formHash){throw 'Form baseline changed'}
$beforeForms=Get-Content -LiteralPath $formPath -Raw|ConvertFrom-Json
$viewsPath=Join-Path $saved.directory 'views.before.json'
if((Get-FileHash -LiteralPath $viewsPath).Hash -cne $saved.viewsHash){throw 'View baseline changed'}
$beforeViews=Get-Content -LiteralPath $viewsPath -Raw|ConvertFrom-Json
Assert-Forms $beforeForms $forms
Assert-Forms $beforeForms (Get-Forms -Draft)
Assert-Views $beforeViews $views
Assert-Views $beforeViews (Get-Views -Draft)
if($Mode -eq 'Apply'){
    if($installed.version -cne $fromVersion -or $solution.version -cne $fromSolutionVersion){throw 'Installed version changed since preparation'}
    $currentBundle=@((Read-Dv 'webresourceset?$select=content&$filter=name eq ''cc_CraigMansell.CMModelDrivenGrid/bundle.js''').value)
    if($currentBundle.Count -ne 1 -or [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Convert]::FromBase64String($currentBundle[0].content))) -cne $fromBundleHash){throw 'Installed baseline asset changed'}
    Assert-Pending (Get-Pending)
    if(!$Package -or $ValidatedPackageHash -notmatch '^[A-Fa-f0-9]{64}$' -or (Get-FileHash -LiteralPath $Package).Hash -cne $ValidatedPackageHash.ToUpperInvariant()){throw 'Exact locally validated package required'}
    $manifest=Read-Package $Package
    $candidateZip=[IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $Package).Path)
    try {
        $expectedEntries=@('solution.xml','customizations.xml','[Content_Types].xml','Controls/cm_CraigMansell.CMModelDrivenGrid/ControlManifest.xml','Controls/cm_CraigMansell.CMModelDrivenGrid/bundle.js','Controls/cm_CraigMansell.CMModelDrivenGrid/bundle.js.LICENSE.txt','Controls/cm_CraigMansell.CMModelDrivenGrid/css/CMModelDrivenGrid.css','Controls/cm_CraigMansell.CMModelDrivenGrid/strings/CMModelDrivenGrid.1033.resx')
        if($candidateZip.Entries.Count -ne $expectedEntries.Count -or @($candidateZip.Entries|Where-Object {$_.FullName -cnotin $expectedEntries}).Count){throw 'Candidate contains unexpected solution components'}
        $reader=[IO.StreamReader]::new($candidateZip.GetEntry('customizations.xml').Open())
        try {[xml]$components=$reader.ReadToEnd()}finally{$reader.Dispose()}
        $customControlNodes=@($components.ImportExportXml.CustomControls.CustomControl)
        if($customControlNodes.Count -ne 1 -or $customControlNodes[0].Name -cne $controlName){throw 'Candidate control scope mismatch'}
        foreach($node in $components.ImportExportXml.ChildNodes){
            if($node.Name -notin @('CustomControls','Languages') -and $node.ChildNodes.Count){throw "Candidate changes outside PCF: $($node.Name)"}
        }
    }finally{$candidateZip.Dispose()}
    if($manifest.controlVersion -cne $version -or $manifest.solutionVersion -cne $solutionVersion){throw 'Candidate version mismatch'}
    $job=[Guid]::NewGuid().ToString()
    Save-Json (Join-Path $saved.directory "$artifactPrefix-import-intent.json") @{target=$target;package=(Resolve-Path -LiteralPath $Package).Path;sha256=$ValidatedPackageHash;importJobId=$job}
    Post-Dv 'ImportSolution' @{CustomizationFile=[Convert]::ToBase64String([IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Package).Path));ImportJobId=$job;PublishWorkflows=$false;OverwriteUnmanagedCustomizations=$false}|Out-Null
    $result=Read-Dv "importjobs($job)?`$select=progress,data,completedon"
    Save-Json (Join-Path $saved.directory "$artifactPrefix-import-result.json") $result
    if($result.progress -ne 100 -or ([xml]$result.data).importexportxml.succeeded -ne 'success'){throw 'Import did not report success'}
    $pending=Get-Pending;Assert-Pending $pending -AllowOwn
    if($pending.webresourceset.Count){
        $nodes=($pending.webresourceset|ForEach-Object {"<webresource>{$($_.webresourceid)}</webresource>"}) -join ''
        Post-Dv 'PublishXml' @{ParameterXml="<importexportxml><webresources>$nodes</webresources></importexportxml>"}|Out-Null
    }
}
if($Mode -in @('Apply','Publish')){
    $intent=Get-Content -LiteralPath (Join-Path $saved.directory "$artifactPrefix-import-intent.json") -Raw|ConvertFrom-Json
    if($intent.target -cne $target -or (Get-FileHash -LiteralPath $intent.package).Hash -cne $intent.sha256){throw 'Imported candidate identity changed'}
    $result=Read-Dv "importjobs($($intent.importJobId))?`$select=progress,data,completedon"
    if($result.progress -ne 100 -or ([xml]$result.data).importexportxml.succeeded -ne 'success'){throw 'Cannot publish an incomplete import'}
    if((Get-Solution).version -cne $solutionVersion -or (Get-Control).version -notin @($fromVersion,$version)){throw 'Unexpected installed versions before publication'}
    if((Get-Control).version -cne $version){
        if(!$AllowPublishAll){throw 'Control manifest needs audited Publish All; no form changes made'}
        $pending=Get-Pending;Assert-Pending $pending -AllowOwn
        Assert-Forms $beforeForms (Get-Forms);Assert-Forms $beforeForms (Get-Forms -Draft)
        Assert-Views $beforeViews (Get-Views);Assert-Views $beforeViews (Get-Views -Draft)
        Assert-Publisher (Get-Publisher)
        Save-Json (Join-Path $saved.directory "$artifactPrefix-pending-before-publish-all.json") $pending
        Post-Dv 'PublishAllXml' @{}|Out-Null
        Save-Json (Join-Path $saved.directory "$artifactPrefix-publish-receipt.json") @{target=$target;time=(Get-Date).ToUniversalTime().ToString('o');reason='Targeted resource publication did not publish the PCF manifest; no unrelated pending UI changes'}
    }
}
$installed=Get-Control;$solution=Get-Solution
if($installed.version -cne $version -or $solution.version -cne $solutionVersion){throw 'Installed version verification failed'}
Assert-Forms $beforeForms (Get-Forms);Assert-Forms $beforeForms (Get-Forms -Draft)
Assert-Views $beforeViews (Get-Views);Assert-Views $beforeViews (Get-Views -Draft)
Assert-Publisher (Get-Publisher)
$intent=Get-Content -LiteralPath (Join-Path $saved.directory "$artifactPrefix-import-intent.json") -Raw|ConvertFrom-Json
if($intent.target -cne $target -or (Get-FileHash -LiteralPath $intent.package).Hash -cne $intent.sha256){throw 'Validated candidate package changed since import'}
$candidate=Read-Package $intent.package
$publishedResources=@((Read-Dv 'webresourceset?$select=webresourceid,name,content&$filter=startswith(name,''cc_CraigMansell.CMModelDrivenGrid/'')').value)
if($publishedResources.Count -ne 3){throw 'Unexpected published PCF resource inventory'}
$verifiedAssets=@()
foreach($resource in $publishedResources){
    $hash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Convert]::FromBase64String($resource.content)))
    if(!$candidate.assets.ContainsKey($resource.name) -or $candidate.assets[$resource.name] -cne $hash){throw "Published resource differs from tested package: $($resource.name)"}
    $verifiedAssets+=@{name=$resource.name;sha256=$hash}
}
Save-Json (Join-Path $saved.directory "$artifactPrefix-published-assets.json") $verifiedAssets
Save-Json (Join-Path $saved.directory "$artifactPrefix-deployment-result.json") @{target=$target;controlVersion=$installed.version;solutionVersion=$solution.version;allItemTemplateFormsUnchanged=$true;allItemTemplateOptionViewsUnchanged=$true;publisherUnchanged=$true;verifiedOn=(Get-Date).ToUniversalTime().ToString('o');runtimeSmokeRequired=$true}
if($Mode -eq 'Export'){
    $result=Export-Pcf (Join-Path $saved.directory 'CMModelDrivenGridSolution.after.unmanaged.zip')
    if($result.controlVersion -cne $version -or $result.solutionVersion -cne $solutionVersion){throw 'Final export version mismatch'}
    Assert-Forms $beforeForms (Get-Forms);Assert-Views $beforeViews (Get-Views)
    Save-Json (Join-Path $saved.directory 'final-package.json') $result
    $releaseAfterPath=Join-Path $saved.directory 'CapricornRelease20260622.after.unmanaged.zip'
    $releaseAfter=Post-Dv 'ExportSolution' @{SolutionName='CapricornRelease20260622';Managed=$false}
    [IO.File]::WriteAllBytes($releaseAfterPath,[Convert]::FromBase64String($releaseAfter.ExportSolutionFile))
    $releaseAfterZip=[IO.Compression.ZipFile]::OpenRead($releaseAfterPath)
    try {
        $releaseAfterReader=[IO.StreamReader]::new($releaseAfterZip.GetEntry('solution.xml').Open())
        try {[xml]$releaseAfterXml=$releaseAfterReader.ReadToEnd()}finally{$releaseAfterReader.Dispose()}
        if($releaseAfterXml.ImportExportXml.SolutionManifest.UniqueName -cne 'CapricornRelease20260622' -or $releaseAfterXml.ImportExportXml.SolutionManifest.Managed -cne '0'){throw 'Final release export unmanaged/identity lock failed'}
    }finally{$releaseAfterZip.Dispose()}
    Assert-Forms $beforeForms (Get-Forms);Assert-Views $beforeViews (Get-Views)
    Save-Json (Join-Path $saved.directory 'final-release-package.json') @{path=$releaseAfterPath;sha256=(Get-FileHash -LiteralPath $releaseAfterPath).Hash;uniqueName='CapricornRelease20260622';managed=$false}
    Write-Output 'Final unmanaged PCF export verified.'
}else{Write-Output 'PCF-only deployment versions and unchanged forms/views verified; signed-in command smoke test still required.'}
