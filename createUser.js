/*
const ADS_UF_DONT_EXPIRE_PASSWD = &H10000;
const ADS_UF_NORMAL = &H200;
const ADS_UF_PASSWD_CANT_CHANGE = &H40;
const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &H6;
const ADS_ACEFLAG_OBJECT_TYPE_PRESENT = &H1;
const CHANGE_PASSWORD_GUID = "{ab721a53-1e2f-11d0-9819-00aa0040529b}";
const ADS_RIGHT_DS_CONTROL_ACCESS = &H100;
*/

function init()		//set up the page on load
	{
	passGen(8);
	accountGroup();	
	}

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
	strGroup = userInfo[7];
	
	strDisplayName = strGiven + " " + strSurname;
	strHomeDrive = "H:";
	strHomeDir = ["\\\\charlie","folders",strUser,"HOME"].join("\\");
	strProfileDir = ["\\\\charlie","folders",strUser,"PROFILE"].join("\\");
	
	strOU = ["ou="+strAccountGroup,"ou="+strAccountType, "ou=Users","ou=Secured Accounts","ou=Secured"].join(",");	//Path to where the user should be created

	//creation and setup
	try
		{
		//Initial setup of environment		
		objRootDSE = GetObject("LDAP://RootDSE");
		strDomain = objRootDSE.Get("DefaultNamingContext");
		objContainer = GetObject("LDAP://" + strOU + "," + strDomain)	//Destination of the user

		//creating user + user information
		objNewUser = objContainer.Create("User", "cn=" + strDisplayName);
		objNewUser.Put("sAMAccountName",strUser);
		objNewUser.Put("UserPrincipalName", strUser + "@girraween-h.schools.win");		
		objNewUser.Put("givenName",strGiven);
		objNewUser.Put("SN",strSurname);
		objNewUser.Put("name",strDisplayName);
		if (strEmail != "")
			{
			objNewUser.Put("mail",strEmail);
			}	
		objNewUser.Put("displayName",strDisplayName);		
		objNewUser.Put("profilePath",strProfileDir);
		objNewUser.Put("homeDrive",strHomeDrive);
		objNewUser.Put("homeDirectory",strHomeDir);		
		//objNewUser.Put("description","Comments");
		objNewUser.Put("password",strPass);
		objNewUser.SetInfo();	//commit base information
		alert("Base info set");
		
		//flags (enable account, password never expires, user can't change password...)
		UAC = objNewUser.get("UserAccountControl");	
		alert("UAC set");
		
		//put user in appropriate group
		objGroup = GetObject("LDAP://cn=" + strGroup + ", ou=Secured Groups, ou=Secured," + strDomain);
		objGroup.add(objNewUser.ADsPath);
		alert(strDisplayName + " added to " + strGroup);
		
		//create folders and grant permissions
				
		
		//cleanup
		objNewUser = null;
		alert("User Created");
		}
	catch(error)	//amazingly useful for debugging, or actually even writing the code!
		{
		alert(error.message);	
		document.main.debug.value = "ERROR: " + error.message;
		}
	}

function accountGroup()	//restrict account_groups based on account types
	{
	studentGroups = new Array(0);	//Auto filled with years
	staffGroups = new Array("Personal","Executive","Shared");
	
	today = new Date();
	year = today.getFullYear()
	
	//year = 12;
	for (i=0;i<6;i++)
		{
		curr = (year - i).toString().substr(2,3);
		studentGroups.push(curr);		
		}
	selection = document.main.account_group;
	selection.length = 0;
	if (document.main.account_type.selectedIndex == 0)	//if student selected
		{
		for (i=0;i<studentGroups.length;i++)
			{
			selection.length += 1;
			selection.options[i] =  new Option(studentGroups[i],studentGroups[i]);
			}
		}
	else
		{
		for (i=0;i<staffGroups.length;i++)
			{
			selection.length += 1;
			selection.options[i] =  new Option(staffGroups[i],staffGroups[i]);
			}
		}
	}

function validate()		//Validation of userinput
	{
	valid = true;						//Assume all's good
	fname = document.main.firstname.value;
	lname = document.main.lastname.value;
	uname = document.main.username.value;
	pass = document.main.password.value;
	if ((fname == "")||(lname == "")||(pass == "" )||(uname == ""))
		{
		valid = false;
		}
	if (valid == true)
		{
		userInfo = pack();
		createUser(userInfo);
		}
	else
		{
		window.alert("Somethin' ain't right...");
		}
	}

function pack()		//Package user information into an array
	{
	//order of user info is:account type,firstname,lastname,username,password,email,groups
	userInfo = new Array(0);	
	userInfo.push(document.main.account_type.value);
	userInfo.push(document.main.account_group.value);
	userInfo.push(document.main.firstname.value);
	userInfo.push(document.main.lastname.value);
	userInfo.push(document.main.username.value);
	userInfo.push(document.main.password.value);
	userInfo.push(document.main.email.value);
	if (userInfo[0] == "Student")
		{
		userInfo.push("Students");
		}
	else
		{
		userInfo.push("Staff");
		}
	return userInfo;	
	}

function passGen(len)	//Generating passwords
	{
	charSet = "abcdefghijklmnopqrstuvwxyz0123456789";
	password = "";
	for (i = 0;i < len;i++)
		{
		password += charSet[(Math.floor(Math.random()*36))];		
		}	
	document.main.password.value = password;
	}
	
function checkUser()
	{
	if (document.main.username.value == "")
		{
		fname = document.main.firstname.value.toLowerCase();
		lname = document.main.lastname.value.toLowerCase();
		username = "";	
		if (fname.length < 6)
			{
			username = fname + lname.substr(0,1);
			}
		else
			{
			username = fname.substr(0,6) + lname.substr(0,1);
			}
		document.main.username.value = username;
		}
	else
		{
		username = document.main.username.value;
		}
	}
