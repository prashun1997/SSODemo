# SSODemo
Office addin with SSO

<b> Steps to be followed <b>
  
  <ol> 
    <li> <h3> Register Azure AD Application </h3> 
    <p> Create a new application under app registrations of your Azure Active Directory. Copy the Application ID and Tenant ID and store it somewhere. Under the authentication tab select both access and ID tokens under the implicit grant flow. Generate a client secret ans dtore the value along with your AppID. Go to the API permissions tab and add the permissions of <i>Files.Read.All, OpenId, profile and User.Read </i>  Grant admin consent to all the permissions. Under the tab expose an API add an Application uri which will of format <i> api:///localhost:3001/app id </i> Now add a scope with the name access_as_user with consent given to both admins and users. Add authorized client applications. The following are the GUIDs of various Office platforms(Office client, office on the web etc) <ul> <li>bc59ab01-8403-45c6-8796-ac3ef710b3e3</li>  <li>57fb890c-0dab-4253-a5e0-7188c88b2bb4</li>  <li>d3590ed6-52b3-4102-aeff-aad2292ab01c</li>  <li>ea5a67f6-b6f3-4338-b240-c655ddc3cc8e</li>  <li>08e18876-6177-487e-b8b5-cf950c1e598c</li> </ul><br>Your application is now registered in Azure Active Directory </p> </li>
      
  <li> <h3> Setting up the project </h3> <p>Use the Startup code folder which contains the inital files for start to creating the demo application. It contains the manifest file for your addin. You just need to modify the <WebApplicationInfo> tag where you need to put your AppId unedr the ID tag and your Application URI under the Resource tag. there is a helper file odata-helper which we will use later for making the graph request. It contains boilerplate code for making a https get request to the graph endpoint. You have inside the public folder index.html that we would be using for our single page application. Nothing fancy just the bare minimum to help you understand the flow. </p> </li>     <li> <h3> Flow of the Demo </h3> <p> I have divided the demo in four sections <ul> <li> <i>Sec 1 - inital setup<i> Contains code for seting up an express app on node js to serve your static index.html page </li>  <li> <i>Sec 2 - id token<i> Contains client side code in index.js file for getting the bootstrap token </li>  <li> <i>Sec 3 - graph token<i> Contains code for fetching access token for graph by the use of bootstrap token acquired earlier </li>  <li> <i>Sec 4 -Final<i> Contains rest of the code which gets one drive items of the logged in user and add it your excel worksheet using the access token for graph acquired earlier </li>  </p> </li>

  </ol>
    
 
    
  <h4><i> Happy Coding ☺️ </i><h4>
