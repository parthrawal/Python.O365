#Created By : Parth Rawal
#Created On: 07-07-2017
#Description: To get the SharePoint Online data

import requests
import json
from sys import argv
import xml.etree.ElementTree as etree  

def SPOnlineHeaders(UserName,Password,TopLevelUrl):
    try:
#Variables Declarations..................................................................
        url="https://login.microsoftonline.com/extSTS.srf"
        headers = {'content-type': 'application/x-www-form-urlencoded'}

        SharepointOnlineAuth="""<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
                    xmlns:a="http://www.w3.org/2005/08/addressing"
                    xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
                <s:Header>
                    <a:Action s:mustUnderstand="1">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>
                    <a:ReplyTo>
                    <a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address>
                    </a:ReplyTo>
                    <a:To s:mustUnderstand="1">https://login.microsoftonline.com/extSTS.srf</a:To>
                    <o:Security s:mustUnderstand="1"
                    xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
                    <o:UsernameToken>
                        <o:Username>{}</o:Username>
                        <o:Password>{}</o:Password>
                    </o:UsernameToken>
                    </o:Security>
                </s:Header>
                <s:Body>
                    <t:RequestSecurityToken xmlns:t="http://schemas.xmlsoap.org/ws/2005/02/trust">
                    <wsp:AppliesTo xmlns:wsp="http://schemas.xmlsoap.org/ws/2004/09/policy">
                        <a:EndpointReference>
                        <a:Address>{}</a:Address>
                        </a:EndpointReference>
                    </wsp:AppliesTo>
                    <t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>
                    <t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>
                    <t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType>
                    </t:RequestSecurityToken>
                </s:Body>
                </s:Envelope>
                """
        SharepointOnlineAuth=SharepointOnlineAuth.format(UserName,Password,TopLevelUrl)
#Let's make call to get the Token in order to get the cookies
        response=requests.post(url,data=SharepointOnlineAuth,headers=headers)
        s=str(response.content)
        Start = [pos for pos in range(len(s)) if s[pos:].startswith('<wsse:BinarySecurityToken Id="Compact0">')][0]
        Finish= [pos for pos in range(len(s)) if s[pos:].startswith('</wsse:BinarySecurityToken>')][0]
        SecurityToken=s[Start+40:Finish]
        Sec_dict={'Authorization':'Bearer'+SecurityToken}
        headers.update(Sec_dict)
#Now, Lets fetch the cookies to make Call to SharePoint Online
        url=TopLevelUrl+'/_forms/default.aspx?wa=wsignin1.0'        
        response=requests.post(url,data=SecurityToken,headers=headers)

#Calling the Sharepoint api Context REST Api using cookies obtained to get X-RequestDigest
        url=TopLevelUrl+"""/_api/Contextinfo"""
        _Fedauth='FedAuth={}'.format(response.cookies['FedAuth'])
        _rtFa='rtFa={}'.format(response.cookies['rtFa'])
        _FinalDict={'Cookie':_Fedauth+';'+_rtFa}
        headers.update(_FinalDict)

        response=requests.post(url,headers=headers)
        s=str(response.content)   

        Start = [pos for pos in range(len(s)) if s[pos:].startswith('<d:FormDigestValue>')][0]
        Finish= [pos for pos in range(len(s)) if s[pos:].startswith('</d:FormDigestValue>')][0]
        Digest=s[Start+19:Finish]
        
        _FinalDict={'X-RequestDigest':Digest}
        headers.update(_FinalDict)        
        _FinalDict={'content-type': 'application/json'}
        headers.update(_FinalDict)
        
    except Exception as e:
        print(str(e))
        raise e

    return headers

def GetDataFromSPOnline(headers,ListName,TopLevelUrl,ServiceUrlWithFilters,path):
    _DataDict={}
    try:
        #Now, Getting the List Information from SharePoint Online    
        url=TopLevelUrl + ServiceUrlWithFilters
        url=url.format(ListName)  
        print(url)      
        _index=0
        ChildDict={}
        response=requests.get(url,headers=headers)
        sxml=str(response.content)
        sxml=sxml[2:len(sxml)-1] #This is to avoid any ' coming in xml
        print(path)
        target = open(path, 'w')
        target.write(sxml)
        tree = etree.parse(path)
        root = tree.getroot()                    
        for child in root:
            if str(child).find('entry') > 0:
                for child1 in child:
                    if str(child1).find('content') > 0:                        
                        ChildDict[str(_index)]={}
                        for child2 in child1[0]:
                            val={child2.tag.split('}')[1]:child2.text}
                            ChildDict[str(_index)].update(val)
                        _DataDict.update(ChildDict)
                        _index+=1
    except Exception as e:
        print('getSPData '+e)
        #raise e
    return _DataDict

def GetAtomFeedDataFromSPOnline(headers,ListName,TopLevelUrl,ServiceUrlWithFilters):
   _DataDict={}
   try:
        #Now, Getting the List Information from SharePoint Online    
        url=TopLevelUrl + ServiceUrlWithFilters
        url=url.format(ListName)  
        print(url)      
        _index=0
        ChildDict={}
        response=requests.get(url,headers=headers)
        sxml=str(response.content)
        sxml=sxml[2:len(sxml)-1] #This is to avoid any ' coming in xml
        tree = etree.fromstring(sxml)
        for child in tree:
            if (str(child).find('entry')>0):
                for child1 in child:
                    if (str(child1).find('content')>0):
                        for child2 in child1:
                            if(str(child2).find('properties')):
                                ChildDict[str(_index)]={}
                                for child3 in child2:
                                    val={child3.tag.replace('{http://schemas.microsoft.com/ado/2007/08/dataservices}',''):child3.text}
                                    ChildDict[str(_index)].update(val)
                                _DataDict.update(ChildDict)
                                _index+=1                           
   except Exception as e:
      print('getSPData '+e)
   return _DataDict

def UpdateDataToSPOnline(headers,TopLevelUrl,ServiceUrlWithFilters,jsonRequest):
    Res='false'
    try:
        #SP.Data.TestcustomFormItem
        #https://bristleconeonline.sharepoint.com/sites/dev/_api/web/lists/GetByTitle('RRF')?$select=ListItemEntityTypeFullName
        #RequestDict={'__metadata':{'type':'SP.Data.TestcustomFormItem'}}
        _newHeader=AdditionalHeaders(headers)
        url=TopLevelUrl+ServiceUrlWithFilters
        print(url)
        response=requests.patch(url,data='{}'.format(jsonRequest),headers=_newHeader)
        Res=str(response)
    except Exception as e:
        Res='false'+e
        raise e
    return Res

def DeleteDataFromSPOnline(headers,TopLevelUrl,ServiceUrlWithFilters,jsonRequest):
    Res='false'
    try:
        _newHeader=AdditionalHeaders(headers)
        url=TopLevelUrl+ServiceUrlWithFilters
        print(url)
        response=requests.delete(url,data='{}'.format(jsonRequest),headers=_newHeader)
        Res=str(response)
    except Exception as e:
        Res='false'+e
        raise e
    return Res

def AdditionalHeaders(_headers):
    _dict={'accept': 'application/json;odata=verbose'}
    _headers.update(_dict)

    _dict={'X-Http-Method': 'PATCH'}
    _headers.update(_dict)

    _dict={'If-Match': '*'}
    _headers.update(_dict)
    return _headers