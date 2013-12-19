Dim IISObj, strQuery, Item

Set IISObj = GetObject("winmgmts://./root/MicrosoftIISv2")
strQuery = "SELECT * FROM IIsWebVirtualDir" 
For Each Item In IISObj.ExecQuery(strQuery)
  WScript.echo Item.Name
Next
