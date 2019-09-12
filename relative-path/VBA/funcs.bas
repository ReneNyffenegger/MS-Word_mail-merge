option explicit



sub setDataSource(doc as document) ' {

'   doc.MailMerge.mainDocumentType = wdFormLetters

    dim xlsx as string

    xlsx = doc.path & "\values.xlsx"

    dim connectString as string
    connectString = "Microsoft.ACE.OLEDB.12.0" & _
                    ";Data Source=" & xlsx

'   connectString = "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" & xlsx _
'   & ";Mode=Read;Extended    Properties=""HDR=YES;IMEX=1;"";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=35;Jet OLEDB:Da"

    doc.mailMerge.openDataSource                                  _
       name                  :=  xlsx                           , _
       sqlStatement          := "SELECT * FROM `sheet_values$`"


'      confirmConversions    :=  false                          , _
'      readOnly              :=  false                          , _
'      format                :=  wdOpenFormatAuto               , _
'      connection            :=  connectString                  , _
'      sqlStatement1         :=  ""                             , _
'      linkToSource          :=  true                           , _
'      addToRecentFiles      :=  false                          , _
'      passwordDocument      :=  ""                             , _
'      passwordTemplate      :=  ""                             , _
'      writePasswordDocument :=  ""                             , _
'      revert                :=  false                          , _
'      subType               :=  wdMergeSubTypeAccess 

'   doc.mailMerge.destination        = wdSendToNewDocument
'   doc.mailMerge.SuppressBlankLines = true


end sub ' }
