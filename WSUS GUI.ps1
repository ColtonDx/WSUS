#########################
#Setup                  #
#########################

#Pull Computer Name from Current PC (Script should be ran fmor WSUS Host)
	$Computer = $env:COMPUTERNAME

#Pull Domain Name from Current PC (must be on same Domain as WSUS)
	$Domain = $env:USERDNSDOMAIN

#Generate Fully Qualified Domain Name of Current PC
	$FQDN = "$Computer" + "." + "$Domain"

#Set the update server to the FQDN
	[String]$updateServer1 = $FQDN

#Disable Secure Connection, enable if this is set in GPO
	[Boolean]$useSecureConnection = $False

#Only change if changed in GPO
	[Int32]$portNumber = 8530

#Load CSV. KBs in this CSV will be declined.
	$path = ""

#Generate a Log File
	$log = "C:\Admin\WSUS\Approved_Updates_{0:MMddyyyy_HHmm}.log" -f (Get-Date)
	new-item -path $log -type file -force

#Load Prerequisites 
	[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
 
 #Connect to Local WSUS
	$updateServer = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($updateServer1,$useSecureConnection,$portNumber)
	$updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
	$u=$updateServer.GetUpdates($updatescope) | ? {($_.IsDeclined -eq $False -AND $_.IsApproved -eq $False)}
 
 #GetGroups
	$wgroups = $updateserver.GetComputerTargetGroups()

##########################
# GUI                    #
##########################

#Main Window
	Add-Type -assembly System.Windows.Forms
	$main_form = New-Object System.Windows.Forms.Form
	$main_form.Text ='Windows Update Launcher'
	$main_form.Width = 600
	$main_form.Height = 400
	$main_form.AutoSize = $true

#Select Group Drop Box
	$Label = New-Object System.Windows.Forms.Label
	$Label.Text = "Select Group to Patch"
	$Label.Location = New-Object System.Drawing.Point(50,10)
	$Label.AutoSize = $true
	$main_form.Controls.Add($Label)
	$GroupComboBox = New-Object System.Windows.Forms.ComboBox
	$GroupComboBox.Width = 300
	Foreach ($wgroup in $wgroups){
		$GroupComboBox.Items.Add($wgroup.Name);
	}
	$GroupComboBox.Location = New-Object System.Drawing.Point(60,40)
	$main_form.Controls.Add($GroupComboBox)
	$setButton = New-Object System.Windows.Forms.Button
	$setButton.Location = New-Object System.Drawing.Size(400,40)
	$setButton.Size = New-Object System.Drawing.Size (120,23)
	$setButton.Text= "Set"
	$setButton.Add_Click({
		$Label2.Text = $GroupComboBox.selectedItem
		$wsusenv = $GroupComboBox.selectedItem
	})
	$main_form.Controls.Add($setButton)

#WSUS Info: Current Host + Domain
	$ComputerLabel = New-Object System.Windows.Forms.Label
	$ComputerLabel.Text = "Current Host: " + $Computer
	$ComputerLabel.Location = New-Object System.Drawing.Point(20,300)
	$ComputerLabel.Autosize = $true
	$main_form.Controls.Add($ComputerLabel)
	$DomainLabel = New-Object System.Windows.Forms.Label
	$DomainLabel.Text = "Domain:  " + $Domain
	$DomainLabel.Location = New-Object System.Drawing.Point(20,320)
	$DomainLabel.Autosize = $true
	$main_form.Controls.Add($DomainLabel)

#Information Labels
	$Label2 = New-Object System.Windows.Forms.Label
	$Label2.Text = ""
	$Label2.Location = New-Object System.Drawing.Point(150,70)
	$Label2.Autosize = $true
	$main_form.Controls.Add($Label2)
	$Label3 = New-Object System.Windows.Forms.Label
	$Label3.Text = "Current Selection:"
	$Label3.Location = New-Object System.Drawing.Point(60,70)
	$Label3.Autosize = $true
	$main_form.Controls.Add($Label3)
   
#Decline ARM Checkbox Label
	$checkboxlabel = New-Object System.Windows.Forms.Label
	$checkboxlabel.Text = "Decline ARM Architecture"
	$checkboxlabel.Location = New-Object System.Drawing.Size(50,110)
	$main_form.Controls.Add($checkboxlabel)
#Decline ARM Checkbox
	$checkbox = New-Object System.Windows.Forms.CheckBox
	$checkbox.Location = New-Object System.Drawing.Size(30,110)
	$main_form.Controls.Add($checkbox)

#Insider Preview Checkbox Label
	$inpcheckboxlabel = New-Object System.Windows.Forms.Label
	$inpcheckboxlabel.Text = "Decline Insider Previews"
	$inpcheckboxlabel.Location = New-Object System.Drawing.Size(50,150)
	$main_form.Controls.Add($inpcheckboxlabel)
#Insider Preview Checkbox
	$inpcheckbox = New-Object System.Windows.Forms.CheckBox
	$inpcheckbox.Location = New-Object System.Drawing.Size(30,150)
	$main_form.Controls.Add($inpcheckbox)

 #Language Checkbox Label
	$langcheckboxlabel = New-Object System.Windows.Forms.Label
	$langcheckboxlabel.Text = "Decline Language Packs"
	$langcheckboxlabel.Location = New-Object System.Drawing.Size(50,190)
	$main_form.Controls.Add($langcheckboxlabel)
#Language Check Box
	$langcheckbox = New-Object System.Windows.Forms.CheckBox
	$langcheckbox.Location = New-Object System.Drawing.Size(30,190)
	$main_form.Controls.Add($langcheckbox)

#Minimum Update Age
	$agecheckboxlabel = New-Object System.Windows.Forms.Label
	$agecheckboxlabel.Text = "Set Minimum Update Age"
	$agecheckboxlabel.Location = New-Object System.Drawing.Size(50,230)
	$agecheckboxlabel.Size = New-Object System.Drawing.Size(70,40)
	$main_form.Controls.Add($agecheckboxlabel)
#Minimum Age Checkbox	
	$agecheckbox = New-Object System.Windows.Forms.CheckBox
	$agecheckbox.Location = New-Object System.Drawing.Size(30,230)
	$agecheckbox.Add_Click({$minagetext.Enabled = $agecheckbox.Checked})
	$main_form.Controls.Add($agecheckbox)
#Minimum Age Text Box	
	$minagetext = New-Object System.Windows.Forms.TextBox
	$minagetext.Location = New-Object System.Drawing.Size(140,232)
	$minagetext.Size = New-Object System.Drawing.Size(50,20)
	$minagetext.Enabled = $agecheckbox.Checked
	$main_form.Controls.Add($minagetext)
#Minimum Age Text Box Label	
	$txtagecheckboxlabel = New-Object System.Windows.Forms.Label
	$txtagecheckboxlabel.Text = "Month(s)"
	$txtagecheckboxlabel.Location = New-Object System.Drawing.Size(190,234)
    $txtagecheckboxlabel.Size = New-Object System.Drawing.Size(50,20)
	$main_form.Controls.Add($txtagecheckboxlabel)
	
#No Approve
	$noapprovelabel = New-Object System.Windows.Forms.Label
	$noapprovelabel.Text = "Approve No Updates"
	$noapprovelabel.Location = New-Object System.Drawing.Size(290,105)
	$noapprovelabel.Size = New-Object System.Drawing.Size(80,40)
	$main_form.Controls.Add($noapprovelabel)
	$noapprovecheckbox = New-Object System.Windows.Forms.CheckBox
	$noapprovecheckbox.Location = New-Object System.Drawing.Size(270,100)
	$noapprovecheckbox.Add_Click({$cusdectext.Enabled = $cusdeccheckbox.Checked})
	$main_form.Controls.Add($noapprovecheckbox)


#Custom Declarations to Filter Out
	$cusdeccheckboxlabel = New-Object System.Windows.Forms.Label
	$cusdeccheckboxlabel.Text = "Declare Custom Filters"
	$cusdeccheckboxlabel.Location = New-Object System.Drawing.Size(290,230)
	$cusdeccheckboxlabel.Size = New-Object System.Drawing.Size(80,40)
	$main_form.Controls.Add($cusdeccheckboxlabel)
	$cusdeccheckbox = New-Object System.Windows.Forms.CheckBox
	$cusdeccheckbox.Location = New-Object System.Drawing.Size(270,230)
	$cusdeccheckbox.Add_Click({$cusdectext.Enabled = $cusdeccheckbox.Checked})
	$main_form.Controls.Add($cusdeccheckbox)
	$cusdectext = New-Object System.Windows.Forms.TextBox
	$cusdectext.Location = New-Object System.Drawing.Size(380,232)
	$cusdectext.Size = New-Object System.Drawing.Size(50,20)
	$cusdectext.Enabled = $cusdeccheckbox.Checked
	$main_form.Controls.Add($cusdectext)


#Patch Button
	$Button1 = New-Object System.Windows.Forms.Button
	$Button1.Location = New-Object System.Drawing.Size(400,310)
	$Button1.Size = New-Object System.Drawing.Size (120,23)
	$Button1.Text= "Patch"
	$Button1.Add_Click({
		if ($agecheckbox.Checked){
			$n = $minagetext.Text
			if ($n -match "^[0-9]+$"){
				BeginPatching ($noapprovecheckbox,$cusdectext,$agecheckbox,$langcheckbox,$cusdeccheckbox,$updateserver,$updatescope,$u,$wsusenv,$updateServer1,$useSecureConnection,$portNumber)
			}
			else{
				$a = new-object -comobject wscript.shell
				$b = $a.popup("Error: Please Only Enter Numbers")
            }
        }
		else{
			BeginPatching ($noapprovecheckbox,$cusdectext,$agecheckbox,$langcheckbox,$cusdeccheckbox,$updateserver,$updatescope,$u,$wsusenv,$updateServer1,$useSecureConnection,$portNumber)
        }
    })
	$main_form.Controls.Add($Button1)

#Load the UI
    $main_form.ShowDialog()


 
#########################
#Function               #
#########################

	function BeginPatching ($noapprovecheckbox,$cusdectext,$agecheckbox,$langcheckbox,$cusdeccheckbox,$updateserver,$updatescope,$u,$wsusenv,$updateServer1,$useSecureConnection,$portNumber){

		#Approved Count
			$ucount = 0
		#Declined Count
			$dcount = 0
       [String]$cusdec = $cusdectext.Text

        #Pull Computer Name from Current PC (Script should be ran fmor WSUS Host)
	        $Computer = $env:COMPUTERNAME

        #Pull Domain Name from Current PC (must be on same Domain as WSUS)
        	$Domain = $env:USERDNSDOMAIN

        #Generate Fully Qualified Domain Name of Current PC
	        $FQDN = "$Computer" + "." + "$Domain"
		#Set the update server to the FQDN
	        [String]$updateServer1 = $FQDN

        #Disable Secure Connection, enable if this is set in GPO
	        [Boolean]$useSecureConnection = $False

        #Only change if changed in GPO
	        [Int32]$portNumber = 8530
        
        #Connect to Local WSUS
	        $updateServer = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($updateServer1,$useSecureConnection,$portNumber)
	        $updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
	        $u=$updateServer.GetUpdates($updatescope) | ? {($_.IsDeclined -eq $False -AND $_.IsApproved -eq $True)}
 			#$n = $minagetext.Text

        [Int32]$n = $minagetext.Text
          $n = (-1*$n)
     
     
     
        #Decline ARM64
		if ($checkbox.Checked){
			foreach ($u1 in $u ){ 
				if ($u1.Title -Like '*ARM64*'){
					#write-host Decline Update : $u1.Title
					#$u1.Decline()
					$dcount=$dcount + 1
				}
			}
		}

	    #Custom Decline Loop
		if ($cusdeccheckbox.Checked){
			foreach ($cusdec in $cusdecs){
				foreach ($u1 in $u){
					if ($u1.Title -Like $cusdec){
                        write-host "Decline" $u1.Title
						#$u1.Decline()
						$dcount=$dcount + 1
					}
				}
			}
		}
	    	
        #Decline LanguagePacks
		if ($langcheckbox.Checked){
		foreach ($u1 in $u ){ 
			if ($u1.Title -Like '*Lang*'){
				#write-host Decline Update : $u1.Title
				#$u1.Decline()
				$dcount=$dcount + 1
				}
			}
		}

        #Decline Insider Previews
		if ($inpcheckbox.Checked){
			foreach ($u1 in $u ){ 
				if ($u1.Title -Like '*Insider*'){
					write-host Decline Update : $u1.Title
					#$u1.Decline()
					$dcount=$dcount + 1
				}
			}
		}
		    #No Approval check
		if ($noapprovecheckbox.Checked){}
			
        else {
             #For Each Approve Loop
		    foreach ($u1 in $u){
		    	if ($checkbox) {
		    		#Only look for Updates older than 6 months
		    		if ($u1.CreationDate -le ((get-date).AddMonths($n))){
		    			#Approve Updates for Current Group
			    		#$u1.Approve("Install", $wgroup)
			    		write-host Approved Update : $u1.Title
			    		$ucount=$ucount + 1
			    	}
			    }
			    else{
			    	#Approve Updates for Current Group
			    	#$u1.Approve("Install", $wgroup)
			    	write-host Approved Update : $u1.Title
			    	#$ucount=$ucount + 1
		    		}
		    	}
		}	
		$a = new-object -comobject wscript.shell
		$b = $a.popup("Approved: " + $ucount + "Declined: " +$dcount)
	}