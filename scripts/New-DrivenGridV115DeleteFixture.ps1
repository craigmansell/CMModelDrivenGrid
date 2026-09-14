[CmdletBinding()]
param([ValidateSet('Create','VerifyDeleted')][string]$Mode='Create')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$dgWorkspace='C:\Users\CraigMansell\Documents\wubl\CrmPackage'
$dgEvidence=Join-Path $dgWorkspace 'outputs/qa/drivengrid-v113-20260914'
. (Join-Path $dgWorkspace 'Migration/DataverseAuth.ps1')
$dgTarget='https://capricornsandbox.crm4.dynamics.com';$dgApi="$dgTarget/api/data/v9.2"
$dgToken=Get-DataverseAccessToken -EnvUrl $dgTarget -Label 'DrivenGrid disposable delete QA'
$dgHeaders=@{Authorization="Bearer $dgToken";Accept='application/json'}
function Read-Dg([string]$Path){Invoke-RestMethod -Uri "$dgApi/$Path" -Headers $dgHeaders -TimeoutSec 45}
if((Read-Dg 'WhoAmI').OrganizationId -ne '27252def-2a75-f111-b27b-000d3ab8db73'){throw 'Sandbox organization lock failed'}
$dgMarker='DGQA20260914:V115-DELETE-ONLY'
$dgIntent=Join-Path $dgEvidence 'v115-delete-fixture-intent.json'
$dgQuery='cm_itemtemplateoptions?$select=cm_itemtemplateoptionid,cm_name,cm_legacyid&$filter=cm_legacyid eq '+[Uri]::EscapeDataString("'$dgMarker'")
if($Mode -eq 'VerifyDeleted'){
    $dgSaved=Get-Content -LiteralPath $dgIntent -Raw|ConvertFrom-Json
    if($dgSaved.target -cne $dgTarget -or $dgSaved.data.cm_legacyid -cne $dgMarker){throw 'Fixture identity changed'}
    if(@((Read-Dg $dgQuery).value).Count -ne 0){throw 'Disposable delete fixture still exists'}
    $dgCheck='cm_itemtemplateoptions?$select=cm_itemtemplateoptionid&$filter=cm_itemtemplateoptionid eq '+$dgSaved.id
    if(@((Read-Dg $dgCheck).value).Count -ne 0){throw 'Original disposable fixture ID still exists'}
    @{target=$dgTarget;id=$dgSaved.id;deleted=$true;verifiedAt=[DateTime]::UtcNow.ToString('o')}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dgEvidence 'v115-delete-fixture-verified.json')
    Write-Output 'Disposable record deletion verified by ID and unique QA marker.';exit
}
if(Test-Path -LiteralPath $dgIntent){throw 'Preserve the existing create intent; do not create a duplicate fixture'}
if(@((Read-Dg $dgQuery).value).Count -ne 0){throw 'QA marker already exists'}
$dgParent=Read-Dg 'cm_itemtemplateoptions(f8a88b32-b515-4b5c-aec0-68b354c2a35b)?$select=cm_legacyid,_cm_itemtemplate_value'
if($dgParent.cm_legacyid -cne 'DGQA20260909:R01' -or $dgParent._cm_itemtemplate_value -cne '79357ab7-61b3-4f6f-bc6b-946fa6a95dae'){throw 'QA parent ownership lock failed'}
$dgId=[Guid]::NewGuid().ToString()
$dgData=@{cm_itemtemplateoptionid=$dgId;cm_name='DrivenGrid QA v115 - delete test only';cm_legacyid=$dgMarker;cm_order=1099;cm_unitcost=1.23;'cm_ItemTemplate@odata.bind'='/cm_itemtemplates(79357ab7-61b3-4f6f-bc6b-946fa6a95dae)';'cm_ParentTemplateOption@odata.bind'='/cm_itemtemplateoptions(f8a88b32-b515-4b5c-aec0-68b354c2a35b)'}
@{target=$dgTarget;id=$dgId;data=$dgData;createdAt=[DateTime]::UtcNow.ToString('o')}|ConvertTo-Json -Depth 10|Set-Content -LiteralPath $dgIntent
Invoke-RestMethod -Uri "$dgApi/cm_itemtemplateoptions" -Headers $dgHeaders -Method Post -ContentType 'application/json' -Body ($dgData|ConvertTo-Json -Compress)|Out-Null
$dgCreated=Read-Dg ("cm_itemtemplateoptions($dgId)"+'?$select=cm_name,cm_legacyid,cm_order,_cm_itemtemplate_value,_cm_parenttemplateoption_value')
if($dgCreated.cm_legacyid -cne $dgMarker -or $dgCreated._cm_itemtemplate_value -cne $dgParent._cm_itemtemplate_value -or $dgCreated._cm_parenttemplateoption_value -cne 'f8a88b32-b515-4b5c-aec0-68b354c2a35b'){throw 'Created fixture scope verification failed'}
$dgCreated|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $dgEvidence 'v115-delete-fixture-created.json')
Write-Output "Created one disposable child for normal native-command deletion: $dgId"
