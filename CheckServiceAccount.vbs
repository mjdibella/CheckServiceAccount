' Kerberos Service Account Checker
' by Michael J. DiBella
' check service accounts for proper KCD operation
' Usage: CheckServiceAccount.vbs dn|upn [account-id=dn|upn]
' Checks current computer account if dn passed without account-id
' Checks current user account if upn passed without account-id
On Error Resume Next 
const VERSION = "v2.0 build 11" 
sMsg = "Kerberos Service Account Checker " & VERSION
if WScript.Arguments.Count > 0 then
	const ADS_MV_LIMIT_WARNING = 1000
	const ADS_NAME_TYPE_DN = 1
	const ADS_NAME_TYPE_USER_PRINCIPAL_NAME = 9
	const ADS_NAME_INITTYPE_GC = 3
	const ADS_NAME_TYPE_NT4 = 3
	const ADS_NAME_TYPE_1779 = 1	
	const ADS_UF_DONT_EXPIRE_PASSWD = &h10000 
	const E_ADS_PROPERTY_NOT_FOUND = &h8000500D 
	const ONE_HUNDRED_NANOSECOND = .000000100 
	const SECONDS_IN_DAY = 86400 
	const EVENTLOG_ERROR = 1
	const EVENTLOG_WARNING = 2
	const EVENTLOG_INFORMATION = 4
	const ADS_UF_LOCKOUT = 16
	const ADS_UF_TRUSTED_FOR_DELEGATION = 524288
	const ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION  = 16777216
	const INVALID_OPTION = "***"
	set rootDSE = GetObject("LDAP://rootDSE")
	iLogSeverity = EVENTLOG_INFORMATION
	if Err = 0 then 
		sDomainDN = Mid("LDAP://" & rootDSE.Get("defaultNamingContext"), 8) 
		set rootDSE = Nothing 
		set objSysInfo = CreateObject("ADSystemInfo") 
		select case LCase(WScript.Arguments.Item(0))
			case "dn"
				if WScript.Arguments.Count > 1 then
					sAccountDN = WScript.Arguments.Item(1)
				else
					set objSystemInfo = CreateObject("ADSystemInfo") 
					set objNetwork = CreateObject("Wscript.Network")
					set objNameTranslate = CreateObject("NameTranslate")
					objNameTranslate.Init ADS_NAME_INITTYPE_GC, ""
					objNameTranslate.Set ADS_NAME_TYPE_NT4, objSystemInfo.DomainShortName & "\" & objNetwork.ComputerName & "$"
					sAccountDN = objNameTranslate.GET(ADS_NAME_TYPE_1779)
					set objSystemInfo = nothing
					set objNetwork = nothing
					set objNameTranslate = nothing
				end if
			case "upn"
				if WScript.Arguments.Count > 1 then
					set objNameTranslate = CreateObject("NameTranslate")
					objNameTranslate.Init ADS_NAME_INITTYPE_GC, ""
					objNameTranslate.set ADS_NAME_TYPE_USER_PRINCIPAL_NAME, WScript.Arguments.Item(1)
					sAccountDn = objNameTranslate.Get(ADS_NAME_TYPE_DN)
					set objNameTranslate = nothing
				else
					set objSysInfo = CreateObject("ADSystemInfo") 
					sAccountDN = objSysInfo.UserName 
					set objSysInfo = nothing
				end if
			case else
				sAccountDn = INVALID_OPTION
		end select
		select case sAccountDN
			case INVALID_OPTION
				sMsg = sMsg & vbCRLF & "The account-id type must be dn or upn.  Correct the account-id type and rerun the utility."
				iLogSeverity = EVENTLOG_WARNING
			case ""
				sMsg = sMsg & vbCRLF & "The account does not exist in directory.  Check the account-id and rerun the utility."
				iLogSeverity = EVENTLOG_WARNING
			case else
				set objAccount = GetObject("LDAP://" & sAccountDN) 
				if objAccount is nothing then
					sMsg = sMsg & vbCRLF & "The account does not exist in directory.  Check the account-id and rerun the utility."
					iLogSeverity = EVENTLOG_WARNING
				else
					sMsg = sMsg & vbCRLF & "Account information for " & sAccountDN & "."
					sMsg = sMsg & vbCRLF & "Display name is " & objAccount.DisplayName & "."
					if objAccount.AccountDisabled then
						sMsg = sMsg & vbCRLF & "The account is disabled. Enable the account and rerun the utility." 
						iLogSeverity = EVENTLOG_ERROR
					else
						sMsg = sMsg & vbCRLF & "The account is enabled." 
					end if
					iUserAccountControl = objAccount.Get("userAccountControl") 
					if (iUserAccountControl and ADS_UF_LOCKOUT) then
						sMsg = sMsg & vbCRLF & "The account is locked out. Unlock and the account and rerun the utility." 
						iLogSeverity = EVENTLOG_ERROR
					else
						sMsg = sMsg & vbCRLF & "The account is not locked." 
					end if
					if ((iUserAccountControl And ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION) = ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION) then
						sMsg = sMsg & vbCRLF & "The account is trusted for constrained delegation to these services:"
						aAllowedToDelegateTo = objAccount.Get("msDS-AllowedToDelegateTo") 
						if IsArray(aAllowedToDelegateTo) then
							for each sSPN in objAccount.Get("msDS-AllowedToDelegateTo") 
								sMsg = sMsg & vbCRLF & "  " & sSPN & " ==> " & GetAccountForSPN(sSPN)
							next
						else
							sSPN = CStr(aAllowedToDelegateTo) 
							sMsg = sMsg & vbCRLF & "  " & " ==> " & GetAccountForSPN(sSPN)
						end if
					else
						sMsg = sMsg & vbCRLF & "The account is not trusted for delegation. Check the delegation settings and rerun the utility." 
						iLogSeverity = EVENTLOG_ERROR
					end if
					iUserCertCount = UBound(objAccount.Get("userCertificate")) + 1
					if iUserCertCount > 0 then
						sMsg = sMsg & vbCRLF & "The account contains " & CStr(iUserCertCount) & " user certificates." 
						if iUserCertCount > ADS_MV_LIMIT_WARNING then
							sMsg = sMsg & vbCRLF & "The number of user certificate is approaching the administrative limit." 
						end if
					end if
					if iUserAccountControl And ADS_UF_DONT_EXPIRE_PASSWD then
					    sMsg = sMsg & vbCRLF & "The password is set not to expire." 
					else 
					    dtmValue = objAccount.PasswordLastChanged 
					    if Err.Number = E_ADS_PROPERTY_NOT_FOUND then
					        sMsg = sMsg & vbCRLF & "The password has never been set. set the password the rerun the utility." 
							if iLogSeverity <> EVENTLOG_ERROR then
								iLogSeverity = EVENTLOG_WARNING
							end if
					    else 
					        intTimeInterval = Int(Now - dtmValue) 
					        sMsg = sMsg & vbCRLF & "The password was last set on " & DateValue(dtmValue) & " at " & TimeValue(dtmValue) & "."
					        sMsg = sMsg & vbCRLF & "The difference between when the password was last set and today is " & intTimeInterval & " days." 
					    end if 
					    set objDomain = GetObject("LDAP://" & sDomainDN) 
					    set objMaxPwdAge = objDomain.Get("maxPwdAge") 
					    if objMaxPwdAge.LowPart = 0 then 
					        sMsg = sMsg & vbCRLF &  "The domain password age policy is set to 0; the password will not expire." 
					    else 
					        dblMaxPwdNano = Abs(objMaxPwdAge.HighPart * 2^32 + objMaxPwdAge.LowPart) 
					        dblMaxPwdSecs = dblMaxPwdNano * ONE_HUNDRED_NANOSECOND
					        dblMaxPwdDays = Int(dblMaxPwdSecs / SECONDS_IN_DAY)
					        sMsg = sMsg & vbCRLF & "Maximum password age for domain " & sDomainDN & " is " & dblMaxPwdDays & " days." 
							iDays = Int((dtmValue + dblMaxPwdDays) - Now)
					        if intTimeInterval >= dblMaxPwdDays then 
					            sMsg = sMsg & vbCRLF & "The password has expired. Reset the password and rerun the utility."
								iLogSeverity = EVENTLOG_ERROR
					        else 
					            sMsg = sMsg & vbCRLF & "The password will expire on " & DateValue(dtmValue + dblMaxPwdDays) & " which is " & iDays & " days from today." 
								sMsg = sMsg & vbCRLF & "Consider setting the password not to expire, or set a reminder to change the password before it expires." 
								if iLogSeverity <> EVENTLOG_ERROR then
									iLogSeverity = EVENTLOG_WARNING
								end if
					        end if 
					    end if 
					end if
				end if
		end select
	else
		sMsg = sMsg & vbCRLF & "A connection to an active directory domain could not be made. Check the machine configuraton and rerun the utility."
		iLogSeverity = EVENTLOG_WARNING
	end if
	select case iLogSeverity
		case EVENTLOG_INFORMATION
			sMsg = sMsg & vbCRLF & "The account is configured correctly."
		case EVENTLOG_WARNING
			sMsg = sMsg & vbCRLF & "The check encountered some warnings. See above for remediation guidance."
		case EVENTLOG_ERROR
			sMsg = sMsg & vbCRLF & "The account configuration will cause errors. See above for remediation guidance."
	end select
	set oShell = WScript.CreateObject("WScript.Shell")
	oShell.LogEvent iLogSeverity, sMsg
else
	sMsg = sMsg & vbCRLF & "Usage: " & wscript.scriptname & " dn|upn [account-id=dn|upn]"
	sMsg = sMsg & vbCRLF & "Checks current computer account if dn passed without account-id."
	sMsg = sMsg & vbCRLF & "Checks current user account if upn passed without account-id."
end if
wscript.echo sMsg

function GetAccountForSPN(sServicePrincipal)
	GetAccountForSPN = ""
	set oADConnection = CreateObject("ADODB.Connection")
	set oADCommand = CreateObject("ADODB.Command")
	oADConnection.Provider = "ADsDSOObject"
	oADConnection.Open "Active Directory Provider"
	set oADCommand.ActiveConnection = oADConnection
	oADCommand.Properties("Page Size") = 1000
	sCommandText = "<LDAP://" & sDomainDN & ">;(servicePrincipalName=" & sServicePrincipal & ");cn;subtree"
	oADCommand.CommandText = sCommandText
	set oResults = oADCommand.Execute
	do until oResults.EOF
		GetAccountForSPN = CStr(oResults.Fields(0).Value)
		oResults.MoveNext
	loop
	if (GetAccountForSPN = "") and left(sServicePrincipal, 5) = "http/" then
		GetAccountForSPN = GetAccountForSPN(Replace(sServicePrincipal, "http/", "HOST/"))
	end if
	set oResults = nothing
	set oADCommand = nothing
	set oADConnection = nothing
end function