[CmdletBinding()]
param(
    [ValidateSet('Snapshot','Oracle','Readback')][string]$Mode='Snapshot',
    [string]$CrmWorkspace='C:\Users\CraigMansell\Documents\wubl\CrmPackage',
    [string]$Evidence='C:\Users\CraigMansell\Documents\wubl\CrmPackage\outputs\qa\drivengrid-v113-20260914'
)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
. (Join-Path $CrmWorkspace 'Migration/DataverseAuth.ps1')
$target='https://capricornsandbox.crm4.dynamics.com';$api="$target/api/data/v9.2"
$token=Get-DataverseAccessToken -EnvUrl $target -Label 'DrivenGrid v113 read-only QA'
$headers=@{Authorization="Bearer $token";Accept='application/json';Prefer='odata.include-annotations="OData.Community.Display.V1.FormattedValue"'}
function Read-Dv([string]$Path){Invoke-RestMethod -Uri "$api/$Path" -Headers $headers -TimeoutSec 60}
function Save-Result([string]$Name,$Value){$path=Join-Path $Evidence $Name;if(Test-Path -LiteralPath $path){throw "Preserve existing QA evidence: $path"};$Value|ConvertTo-Json -Depth 80|Set-Content -LiteralPath $path -Encoding utf8}
if((Read-Dv 'WhoAmI').OrganizationId -ne '27252def-2a75-f111-b27b-000d3ab8db73'){throw 'Sandbox organization lock failed'}
New-Item -ItemType Directory -Path $Evidence -Force|Out-Null
$template='79357ab7-61b3-4f6f-bc6b-946fa6a95dae';$parent='f8a88b32-b515-4b5c-aec0-68b354c2a35b'
$fixtures=@(Get-Content -LiteralPath (Join-Path $CrmWorkspace 'outputs/deployment-backups/drivengrid-full-qa-sandbox-20260909-110800/fixtures.verified.json') -Raw|ConvertFrom-Json)
if($fixtures.Count -ne 107 -or @($fixtures|Where-Object {$_.key -eq 'template' -and $_.id -eq $template}).Count -ne 1){throw 'QA fixture identity changed'}
if($Mode -in @('Snapshot','Readback')){
    $sets=@{cm_itemtemplateoptions='cm_itemtemplateoptionid';cm_itemtemplates='cm_itemtemplateid';cm_componentoptions='cm_componentoptionid'}
    $records=[Collections.Generic.List[object]]::new()
    foreach($set in $sets.Keys){
        $members=@($fixtures|Where-Object set -eq $set)
        for($offset=0;$offset -lt $members.Count;$offset+=30){
            $batch=@($members|Select-Object -Skip $offset -First 30)
            $filter=($batch|ForEach-Object {"$($sets[$set]) eq $($_.id)"}) -join ' or '
            $response=Read-Dv ($set+'?$filter='+[Uri]::EscapeDataString($filter))
            if($response.PSObject.Properties.Name -contains '@odata.nextLink'){throw 'Snapshot pagination required'}
            foreach($fixture in $batch){
                $found=@($response.value|Where-Object {$_.$($sets[$set]) -eq $fixture.id})
                if($found.Count -eq 0 -and $fixture.key -eq 'C65'){continue}
                if($found.Count -ne 1){throw "Missing/duplicate QA fixture $($fixture.key)"}
                if($set -eq 'cm_itemtemplateoptions' -and $found[0].cm_legacyid -cne $fixture.data.cm_legacyid){throw 'Fixture ownership marker changed'}
                $records.Add(@{key=$fixture.key;id=$fixture.id;set=$set;data=$found[0]})
            }
        }
    }
    if($records.Count -ne 106){throw "QA inventory changed: $($records.Count)"}
    if($Mode -eq 'Snapshot'){
        Save-Result 'records.before.json' @($records)
        Write-Output "Snapshotted $($records.Count) existing QA records; historical C65 deletion preserved."
    }else{
        $before=Get-Content -LiteralPath (Join-Path $Evidence 'records.before.json') -Raw|ConvertFrom-Json
        $changes=[Collections.Generic.List[object]]::new()
        foreach($record in $records){
            $original=@($before|Where-Object id -eq $record.id)[0]
            foreach($property in $original.data.PSObject.Properties){
                if($property.Name -match '@|^modifiedon$|^versionnumber$'){continue}
                $current=$record.data.PSObject.Properties[$property.Name]
                $value=if($current){$current.Value}else{$null}
                if(($property.Value|ConvertTo-Json -Compress -Depth 10) -cne ($value|ConvertTo-Json -Compress -Depth 10)){$changes.Add(@{key=$record.key;id=$record.id;field=$property.Name;before=$property.Value;after=$value})}
            }
        }
        Save-Result 'records.after.json' @($records)
        Save-Result 'record-audit.json' @{target=$target;checkedAt=[DateTime]::UtcNow.ToString('o');records=$records.Count;changes=@($changes)}
        Write-Output "QA readback: $($records.Count) records, $($changes.Count) changed business fields."
        if($changes.Count){throw 'QA values have not all been restored; inspect record-audit.json'}
    }
    exit
}
function Esc($Value){[Security.SecurityElement]::Escape([string]$Value)}
function Condition($Field,$Operator,$Value=$null){$valueXml=if($null -ne $Value){" value='$(Esc $Value)'"}else{''};"<condition attribute='$Field' operator='$Operator'$valueXml/>"}
$cases=[Collections.Generic.List[object]]::new()
function Add-Case($Column,$Field,$Label,$Operator,$InputKind='none',$Value=$null,$Upper=$null){
    $predicate=if($InputKind -eq 'range'){(Condition $Field 'ge' $Value)+(Condition $Field 'le' $Upper)}elseif($InputKind -eq 'date-range'){(Condition $Field 'on-or-after' $Value)+(Condition $Field 'on-or-before' $Upper)}elseif($InputKind -eq 'periodYear'){"<condition attribute='$Field' operator='$Operator'><value>$Value</value><value>$Upper</value></condition>"}else{Condition $Field $Operator $Value}
    $cases.Add(@{column=$Column;field=$Field;label=$Label;operator=$Operator;input=$InputKind;value=$Value;upper=$Upper;predicate=$predicate})
}
$text=@(@('Equals','eq','Alpha red'),@('Does not equal','ne','Alpha red'),@('Contains data','not-null'),@('Does not contain data','null'),@('Contains','like','%red%'),@('Does not contain','not-like','%red%'),@('Begins with','like','Alpha%'),@('Does not begin with','not-like','Alpha%'),@('Ends with','like','%red'),@('Does not end with','not-like','%red'))
foreach($c in $text){Add-Case 'Name' 'cm_name' $c[0] $c[1] $(if($c.Count -gt 2){'value'}else{'none'}) $(if($c.Count -gt 2){$c[2]}else{$null})}
foreach($pair in @(@('apostrophe',"O'Brien & Sons"),@('no-match','No such QA value'))){Add-Case 'Name' 'cm_name' 'Equals' 'eq' 'value' $pair[1]}
foreach($literal in @('%','_','[')){Add-Case 'Name' 'cm_name' 'Contains' 'like' 'value' ('%'+$(if($literal -eq '['){'[[]'}else{"[$literal]"})+'%')}
$numeric=@(@('Equals','eq','10'),@('Does not equal','ne','10'),@('Contains data','not-null'),@('Does not contain data','null'),@('Greater than','gt','10'),@('Greater than or equal to','ge','10'),@('Less than','lt','10'),@('Less than or equal to','le','10'),@('Between','range','10','50'))
foreach($c in $numeric){Add-Case 'Unit Cost' 'cm_unitcost' $c[0] $c[1] $(if($c.Count -gt 3){'range'}elseif($c.Count -gt 2){'value'}else{'none'}) $(if($c.Count -gt 2){$c[2]}else{$null}) $(if($c.Count -gt 3){$c[3]}else{$null})}
foreach($spec in @(@('Option','cm_option','55559ecf-9ece-42d5-a414-ab3c42aca7a4'),@('Status Reason','statuscode','2'))){foreach($c in @(@('Equals','eq'),@('Does not equal','ne'),@('Contains data','not-null'),@('Does not contain data','null'))){Add-Case $spec[0] $spec[1] $c[0] $c[1] $(if($c[1] -in @('eq','ne')){'choice'}else{'none'}) $(if($c[1] -in @('eq','ne')){$spec[2]}else{$null})}}
$dates=@(@('Contains data','not-null'),@('Does not contain data','null'),@('After','gt','2026-09-01'),@('Before','lt','2026-09-01'),@('Between','range','2026-09-01','2026-09-08'),@('On','on','2026-09-08'),@('On or before','on-or-before','2026-09-01'),@('On or after','on-or-after','2026-09-01'),@('Yesterday','yesterday'),@('Today','today'),@('Tomorrow','tomorrow'),@('Last 7 days','last-seven-days'),@('Next 7 days','next-seven-days'),@('Last week','last-week'),@('This week','this-week'),@('Last month','last-month'),@('This month','this-month'),@('Last year','last-year'),@('This year','this-year'),@('Last X days','last-x-days','10'),@('Next X days','next-x-days','10'),@('Last X months','last-x-months','2'),@('Next X months','next-x-months','2'))
foreach($c in $dates){Add-Case 'Created On' 'createdon' $c[0] $c[1] $(if($c.Count -gt 3){'date-range'}elseif($c[1] -match '-x-'){'count'}elseif($c.Count -gt 2){'date'}else{'none'}) $(if($c.Count -gt 2){$c[2]}else{$null}) $(if($c.Count -gt 3){$c[3]}else{$null})}
$extra=@(@('Next week','next-week'),@('Next month','next-month'),@('Next year','next-year'),@('Next X hours','next-x-hours','3'),@('Last X hours','last-x-hours','3'),@('Next X weeks','next-x-weeks','3'),@('Last X weeks','last-x-weeks','3'),@('Next X years','next-x-years','3'),@('Last X years','last-x-years','3'),@('This fiscal year','this-fiscal-year'),@('This fiscal period','this-fiscal-period'),@('Next fiscal year','next-fiscal-year'),@('Next fiscal period','next-fiscal-period'),@('Last fiscal year','last-fiscal-year'),@('Last fiscal period','last-fiscal-period'),@('Last X fiscal years','last-x-fiscal-years','3'),@('Last X fiscal periods','last-x-fiscal-periods','3'),@('Next X fiscal years','next-x-fiscal-years','3'),@('Next X fiscal periods','next-x-fiscal-periods','3'),@('In fiscal year','in-fiscal-year','2026'),@('In fiscal period','in-fiscal-period','9'),@('In fiscal period and year','in-fiscal-period-and-year','9','2026'),@('In or before fiscal period and year','in-or-before-fiscal-period-and-year','9','2026'),@('In or after fiscal period and year','in-or-after-fiscal-period-and-year','9','2026'),@('Older than X months','olderthan-x-months','3'),@('Older than X years','olderthan-x-years','3'),@('Older than X weeks','olderthan-x-weeks','3'),@('Older than X days','olderthan-x-days','3'),@('Older than X hours','olderthan-x-hours','3'),@('Older than X minutes','olderthan-x-minutes','3'))
foreach($c in $extra){Add-Case 'Created On' 'createdon' $c[0] $c[1] $(if($c.Count -gt 3){'periodYear'}elseif($c.Count -gt 2){'count'}else{'none'}) $(if($c.Count -gt 2){$c[2]}else{$null}) $(if($c.Count -gt 3){$c[3]}else{$null})}
foreach($c in @(@('Contains data','not-null'),@('Does not contain data','null'))){Add-Case 'Record Created On' 'overriddencreatedon' $c[0] $c[1]}
$results=[Collections.Generic.List[object]]::new()
foreach($scope in @('root','child')){
    foreach($case in $cases){
        $parentFilter=if($scope -eq 'root'){Condition 'cm_parenttemplateoption' 'null'}else{Condition 'cm_parenttemplateoption' 'eq' $parent}
        $fetch="<fetch mapping='logical' count='5000'><entity name='cm_itemtemplateoption'><attribute name='cm_itemtemplateoptionid'/><attribute name='cm_legacyid'/><order attribute='cm_order'/><order attribute='cm_itemtemplateoptionid'/><filter type='and'>$(Condition 'cm_itemtemplate' 'eq' $template)$parentFilter$($case.predicate)</filter></entity></fetch>"
        try{
            $response=Read-Dv ('cm_itemtemplateoptions?fetchXml='+[Uri]::EscapeDataString($fetch))
            if($response.PSObject.Properties.Name -contains '@odata.nextLink'){throw 'Oracle pagination required'}
            $results.Add(@{scope=$scope;case=$case;accepted=$true;keys=@($response.value|ForEach-Object {($_.cm_legacyid -split ':')[-1]});ids=@($response.value|ForEach-Object {$_.cm_itemtemplateoptionid})})
        }catch{$results.Add(@{scope=$scope;case=$case;accepted=$false;error=$_.Exception.Message})}
    }
}
Save-Result 'filter-oracle.validated.json' @{target=$target;template=$template;createdAt=[DateTime]::UtcNow.ToString('o');results=@($results)}
Write-Output "Oracle queries: $($results.Count), rejected: $(@($results|Where-Object accepted -eq $false).Count)."
