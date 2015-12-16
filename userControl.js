//functions that will allow changes to be made to user account(s)
var selected = new Array(0);	//will store actual selected users
function searchUser(strName)
	{
	var arrResult = [];
	
	objRootDSE = GetObject("LDAP://RootDSE");
	strDomain = objRootDSE.Get("DefaultNamingContext");

	strOU = "OU=Student,OU=Users,OU=Secured,OU=Secured Accounts"; 	// Set the OU to search here.
	strAttrib = "name,samaccountname,password"; 			// Set the attributes to retrieve here.

	objConnection = new ActiveXObject("ADODB.Connection");	//create connection to AD
	objConnection.Provider="ADsDSOObject";
	objConnection.Open("ADs Provider");
	objCommand = new ActiveXObject("ADODB.Command");		//command object
	objCommand.ActiveConnection = objConnection;
	Domain = "LDAP://"+strOU+","+strDomain;
	arrAttrib = strAttrib.split(",");
	
	//sql query for AD
	objCommand.CommandText = "SELECT '"+strAttrib+"' from '"+Domain+"' WHERE objectCategory = 'user' AND objectClass='user' AND samaccountname='"+strName+"*' ORDER BY samaccountname ASC";
	
	try 
		{
		objRecordSet = objCommand.Execute();	//execute the command, retrieve raw data 

		objRecordSet.Movefirst;
		while(!(objRecordSet.EoF)) 				//go through and extract data and append it to arrResult
			{
			var locarray = new Array(0);
			for(y = 0; y < arrAttrib.length; y++) 
				{
				locarray.push(objRecordSet.Fields[y].value);
				}
			arrResult.push(locarray);
			objRecordSet.MoveNext;
			}
		return arrResult;
		} 
	catch(error) 
		{
		alert(error.message);
		}
	}

function getUsers(entry)
	{	
	//document.main.s1.options[].value = userList;
	//var userList;
	users = new Array("Jack","Jason","Jester","Jill","Jim","Jones","John");
	collection = new Array(0);	//collection will store captured users
	lenEntry = entry.length;
	for (i=0;i<users.length;i++)
		{
		currUser = users[i].toLowerCase();
		if (entry == currUser.substr(0,lenEntry))
			{			
			collection.push(currUser);
			}
		}
	return collection;
	}

function updateList()
	{
	//document.main.debug.value = document.main.userlike.value;
	entry = document.main.userlike.value.toLowerCase();
	userList = getUsers(entry);		//gets potential users based on 'entry'
	selection = document.main.retrieved;	//handle to control the 'retrieved' form
	
	selection.length = 0;			//reset selection
	if ((userList.length == 0)||(entry == ""))
		{		
		selection.length += 1;
		selection.options[0] = new Option("None");	//just neater is all
		selection.disabled = true;					//So can't add "None" users
		document.main.debug.value = "None";			
		}
	else 
		{
		selection.disabled = false;					
		document.main.debug.value = userList;			
		
		//populate 'retrieved' with users
		for (i=0; i<userList.length; i++)
			{
			selection.length += 1;					//create space for new option
			selection.options[i] = new Option(userList[i],userList[i],false,false);	//add each user
			}
		}
	}

function addUsers()	//goes through selected users in retrieved and adds them
	{
	selection = document.main.retrieved;
	n = 0;
	for (i=0;i<selection.length;i++)
		{
		if (selection[i].selected == true);
			{
			n += 1;			
			selected.push(selection[i].value);
			}
		}
	document.main.selected.value = selected;
	}