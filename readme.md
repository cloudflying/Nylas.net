#Nylas .net class library

**Initial Setup**

1.    Reference the Nylas.net Library    
	     `Imports NylasApp`
2.    Setup API properties    
	`nylas.nylasClientID = "<YOUR NYLAS CLIENT ID>" `
        `nylas.nylasClientSecret = "<YOUR NYLAS CLIENT SECRET>" `
        `nylas.redirectAddress = "<OAUTH REDIRECT URL - MAKE SURE THIS MATCHES ENTRY ON NYLAS DASHBOARD>" `
3.    Call OAuth Flow first.   



**OAuth Flow**   

Initial Page   
    response.redirect(nylas.oauthredirect)
    '// Include User Email Address if you know it.
    response.redirect(nylas.oauthredirect("john@doe.net"))


Redirect Page
     '// Exchange OAuth Code for Token - Store this token for all future requests.
     Dim token As String = nylas.oauthTokenExchange(Request.QueryString("code"))




**Retrieve ALL messages**   
API Documentation : (https://www.nylas.com/docs/platform#messages)    

    Try
            Dim lst As New List(Of messageObject)
            lst = nylas.retrieveAllMessages("<TOKEN FROM OAUTH FLOW>")
           
            For Each msg As messageObject In lst
				' Do whatever you want with the message list
				
            Next
    Catch ex As Exception
            MsgBox(ex.Message)
    End Try

		
**Release Notes**
This release is very early, it is not feature complete and is constantly being updated. Check back to make sure you have the most recent version. 

**Requirements**
You must have JSON.net included in your project.