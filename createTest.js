function createUser(userInfo)
	{
	//account info setup
	strAccountType = userInfo[0];
	strAccountGroup = userInfo[1];
	strGiven = userInfo[2];
	strSurname = userInfo[3];
	strUser = userInfo[4];
	strPass = userInfo[5];
	strEmail = userInfo[6];
	groups = userInfo[7];
	
	strDisplayName = strGiven + " " + strSurname;
	strHomeDrive = "H:";
	strHomeDir = ["\\charlie","folders",strUser,"HOME"].join("\\");
	strProfileDir = ["\\charlie","folders",strUser,"PROFILE"].join("\\");
	
	//Initial setup of environment
	strOU = ["ou="+strAccountGroup+,"ou="+strAccountType+"ou=Users","ou=Secured Accounts","ou=Secured"].join(",")	//Path to where the user should be created
	objRootDSE = GetObject("LDAP://RootDSE");
	strDomain = objRootDSE.Get("DefaultNamingContext");
	objContainer = GetObject("LDAP://" + strOU + "," + strDomain)	//Destination of the user

	WScript.echo(strDomain);
	//creation and setup
	try
		{
		objNewUser = objContainer.Create("User", "cn=" + strDisplayName);
		objNewUser.Put("sAMAccountName",strUser);
		objNewUser.Put("UserPrincipalName",strUser + "@girraween-h.schools.win");		
		objNewUser.Put("givenName",strGiven);
		objNewUser.Put("SN",strSurname);
		if (strEmail != "")
			{
			objNewUser.Put("mail",strEmail);
			}	
		objNewUser.Put("displayName",strDisplayName);
		objNewUser.Put("profilePath",strProfileDir);
		objNewUser.Put("homeDrive",strHomeDrive);
		objNewUser.Put("homeDirectory",strHomeDir);		
		objNewUser.Put("description","Comments");
		objNewUser.Put("password",strPass);
		//objNewUser.Properties["userAccountControl"].Value = 0x200; 
		UAC = objNewUser.UserAccountControl;
		objNewUser.SetInfo();	
		}
	catch(error)
		{
		WScript.echo(error.message);
		}
	}
userInfo = ['Student','07','foo','bar','foob','43jf03j','soomba','Students'];
createUser(userInfo);