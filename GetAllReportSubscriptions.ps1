# Parameters:
#    currentOwner - DOMAIN\USER that owns the subscriptions you wish to change
#    server        - server and instance name (e.g. myserver/reportserver or myserver/reportserver_db2)
#    site        - use "/" for default native mode site


#		Example of connecting to a SharePoint 2013 Report server. 
#    .\GetAllReportSubscriptions.ps1  "<ServerURL>/<ReportSiteNameURL>/_vti_bin/ReportServer/" "http://<ServerUrl>/<ReportSiteNameURL>"
#   Example of SSRS stand along mode
#    .\GetAllReportSubscriptions.ps1 "<ReportServer>/Reportserver/" "/<ReportFolder>"

Param(
    
    [string]$server,
    [string]$site
)

$ReportData=@()


$rs2010 = New-WebServiceProxy -Uri "http://$server/ReportService2010.asmx" -Namespace SSRS.ReportingService2010 -UseDefaultCredential ;
$subscriptions += $rs2010.ListSubscriptions($site); 

Write-Host " "
Write-Host " "
Write-Host "----- Reports: "



foreach ( $rpt in $subscriptions) {
	write-Host $rpt.report
	$item=@{}
	$item.ReportName =  $rpt.report
	$item.Owner = $rpt.Owner
	$item.lastexecuted = $rpt.lastexecuted
	$type = $rs2010.GetType().Namespace
	$ExtensionSettingsDataType = ($type + ".ExtensionSettings")
		$ExtensionSettingsObj = New-Object ($ExtensionSettingsDataType)
		$ExtensionSettingsObj = $null
		$Description = $null
		$ActiveStateObj = $null
		$Status = $null
		$EventType = $null
		$MatchData = $null
		$ParameterValueObj = $null
		$ExtensionSettingsObj
	$x = $rs2010.GetSubscriptionProperties($rpt.SubscriptionID, `
                    
                    
                    [ref]$ExtensionSettingsObj, `
                                        [ref]$Description, `
                                        [ref]$ActiveStateObj, `
                                        [ref]$Status, `
                                        [ref]$EventType, `
                                        [ref]$MatchData, `
                                        [ref]$ParameterValueObj );
		
		$item.ScheduleData = $MatchData
		foreach ($param in $ExtensionSettingsObj.ParameterValues){
			#write-host $param.Name
			if($param.Name -eq "TO"){
				$item.TO = $param.Value
			}
			if($param.Name -eq "CC"){
				$item.CC = $param.Value
			}
			if($param.Name -eq "Subject"){
				$item.Subject = $param.Value
			}
			if($param.Name -eq "Comment"){
				$item.Comment = $param.Value
			}
			if($param.Name -eq "RenderFormat"){
				$item.RenderFormat = $param.Value
			}
		 }
		 $rptPrms = ""
		 foreach ($param in $ParameterValueObj){
			#write-host $param.Name  " = "  $param.Value  
			if($param.Value -ne "" -and $param.Value -ne $null){
				
				$rptPrms += $param.Name + " = " + $param.Value + "," |echo
			}
		 }
		 $item.ReportParameters  = $rptPrms
		 
	$rptItem = new-object -TypeName PSObject -prop $item
	$rptItem = $rptItem |select-object ReportName,TO,CC,Subject,Comment,lastexecuted,Owner,RenderFormat,ReportParameters,ScheduleData
	
	
	$ReportData+=$rptItem
}


$nowtime = Get-Date -format yyyyMMddHHmmss
$filenamecsv = "SSRS_ReportSubscriptions_$nowtime.csv"
Write-Host " "
$ReportData | Export-Csv -notype  $filenamecsv
write-host $filenamecsv
