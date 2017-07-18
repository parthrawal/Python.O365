# SharePointOnline
This py file can help Authenticate SharePoint Online data and then supporting methods can help fetch records from List

Add this File to your solution 

import sys
sys.path.insert(0,'C:\<foldername>\PythonFiles')
import SharePointOnline

After importing use following code to get the required headers for accessing the data on SharePoint Online.

_headers=SharePointOnline.SPOnlineHeaders("User Email Address","Password","https://XXX-YYY.sharepoint.com")    
    _DataDict=SharePointOnline.GetAtomFeedDataFromSPOnline(_headers,"ListName to access","https://XXX-YYY.sharepoint.com","/sites/dev/_api/web/lists/getByTitle('{}')/Items")
    
Above code should return a dictionary with required data.

I have checked with Python 3, it is working just fine.

Cheers!
    

