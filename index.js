var path = require('path'), 
    moment = require('moment'),
    crypto = require('crypto'),
    xml2js = require('xml2js')
  ;


exports.client = null;

exports.initialize = function(settings, callback) {
  var soap = require('soap');
  // TODO: Handle different locations of where the asmx lives.
  var endpoint = 'https://' + path.join(settings.url, 'EWS/Exchange.asmx');
  var url = path.join(__dirname, 'Services.wsdl');

  soap.createClient(url, {}, function(err, client) {
    if (err) {
      return callback(err);
    }
    if (!client) {
      return callback(new Error('Could not create client'));
    }

    exports.client = client;
    exports.client.setSecurity(new soap.BasicAuthSecurity(settings.username, settings.password));
    exports.client.addSoapHeader('<t:RequestServerVersion Version="Exchange2010" />');
    exports.client.addSoapHeader('<t:TimeZoneContext> <t:TimeZoneDefinition Id="Central Standard Time" /> </t:TimeZoneContext>');

    // return callback(null);
    return callback(settings);
  }, endpoint);
};


exports.getEmails = function(folderName, limit, callback) {
  if (typeof(folderName) === "function") {
    callback = folderName;
    folderName = 'inbox';
    limit = 10;
  }
  if (typeof(limit) === "function") {
    callback = limit;
    limit = 10;
  }
  if (!exports.client) {
    return callback(new Error('Call initialize()'));
  }

  var soapRequest = 
    '<tns:FindItem Traversal="Shallow" xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
      '<tns:ItemShape>' +
        '<t:BaseShape>IdOnly</t:BaseShape>' +
        '<t:AdditionalProperties>' +
          '<t:FieldURI FieldURI="item:ItemId"></t:FieldURI>' +
          // '<t:FieldURI FieldURI="item:ConversationId"></t:FieldURI>' +
          // '<t:FieldURI FieldURI="message:ReplyTo"></t:FieldURI>' +
          // '<t:FieldURI FieldURI="message:ToRecipients"></t:FieldURI>' +
          // '<t:FieldURI FieldURI="message:CcRecipients"></t:FieldURI>' +
          // '<t:FieldURI FieldURI="message:BccRecipients"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:DateTimeCreated"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:DateTimeSent"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:HasAttachments"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:Size"></t:FieldURI>' +
          '<t:FieldURI FieldURI="message:From"></t:FieldURI>' +
          '<t:FieldURI FieldURI="message:IsRead"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:Importance"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:Subject"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:DateTimeReceived"></t:FieldURI>' +
        '</t:AdditionalProperties>' + 
      '</tns:ItemShape>' +
      '<tns:IndexedPageItemView BasePoint="Beginning" Offset="0" MaxEntriesReturned="10"></tns:IndexedPageItemView>' +
      '<tns:ParentFolderIds>' + 
        '<t:DistinguishedFolderId Id="inbox"></t:DistinguishedFolderId>' + 
      '</tns:ParentFolderIds>' + 
    '</tns:FindItem>';

  exports.client.FindItem(soapRequest, function(err, result, body) {
    if (err) {
      return callback(err);
    }

    var parser = new xml2js.Parser();

    parser.parseString(body, function(err, result) {
      var responseCode = result['s:Body']['m:FindItemResponse']['m:ResponseMessages']['m:FindItemResponseMessage']['m:ResponseCode'];

      if (responseCode !== 'NoError') {
        return callback(new Error(responseCode));
      }
        
      var rootFolder = result['s:Body']['m:FindItemResponse']['m:ResponseMessages']['m:FindItemResponseMessage']['m:RootFolder'];
      
      var emails = [];
      rootFolder['t:Items']['t:Message'].forEach(function(item, idx) {
        var md5hasher = crypto.createHash('md5');
        md5hasher.update(item['t:Subject'] + item['t:DateTimeSent']);
        var hash = md5hasher.digest('hex');

        var itemId = {
          id: item['t:ItemId']['@'].Id,
          changeKey: item['t:ItemId']['@'].ChangeKey
        };

        var dateTimeReceived = item['t:DateTimeReceived'];

        emails.push({
          id: itemId.id + '|' + itemId.changeKey,
          hash: hash,
          subject: item['t:Subject'],
          dateTimeReceived: moment(dateTimeReceived).format("MM/DD/YYYY, h:mm:ss A"),
          size: item['t:Size'],
          importance: item['t:Importance'],
          hasAttachments: (item['t:HasAttachments'] === 'true'),
          from: item['t:From']['t:Mailbox']['t:Name'],
          isRead: (item['t:IsRead'] === 'true'),
          meta: {
            itemId: itemId
          }
        });
      });

      callback(null, emails);
    });
  });
};


exports.getEmail = function(itemId, callback) {
  if (!exports.client) {
    return callback(new Error('Call initialize()'))
  }
  if ((!itemId['id'] || !itemId['changeKey']) && itemId.indexOf('|') > 0) {
    var s = itemId.split('|');

    itemId = {
      id: itemId.split('|')[0],
      changeKey: itemId.split('|')[1]
    };
  }

  if (!itemId.id || !itemId.changeKey) {
    return callback(new Error('Id is not correct.'));
  }

  var soapRequest = 
    '<tns:GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
      '<tns:ItemShape>' +
        '<t:BaseShape>Default</t:BaseShape>' +
        '<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
      '</tns:ItemShape>' +
      '<tns:ItemIds>' +
        '<t:ItemId Id="' + itemId.id + '" ChangeKey="' + itemId.changeKey + '" />' +
      '</tns:ItemIds>' +
    '</tns:GetItem>';

  exports.client.GetItem(soapRequest, function(err, result, body) {
    if (err) {
      return callback(err);
    }

    var parser = new xml2js.Parser();

    parser.parseString(body, function(err, result) {
      var responseCode = result['s:Body']['m:GetItemResponse']['m:ResponseMessages']['m:GetItemResponseMessage']['m:ResponseCode'];

      if (responseCode !== 'NoError') {
        return callback(new Error(responseCode));
      }
       
      var item = result['s:Body']['m:GetItemResponse']['m:ResponseMessages']['m:GetItemResponseMessage']['m:Items']['t:Message'];

      var itemId = {
        id: item['t:ItemId']['@'].Id,
        changeKey: item['t:ItemId']['@'].ChangeKey
      };

      function handleMailbox(mailbox) {
        var mailboxes = [];

        if (!mailbox || !mailbox['t:Mailbox']) {
          return mailboxes;
        }
        mailbox = mailbox['t:Mailbox'];

        function getMailboxObj(mailboxItem) {
          return {
            name: mailboxItem['t:Name'],
            emailAddress: mailboxItem['t:EmailAddress']
          };
        }

        if (mailbox instanceof Array) {
          mailbox.forEach(function(m, idx) {
            mailboxes.push(getMailboxObj(m));
          })
        } else {
          mailboxes.push(getMailboxObj(mailbox));
        }

        return mailboxes;
      }

      var toRecipients = handleMailbox(item['t:ToRecipients']);
      var ccRecipients = handleMailbox(item['t:CcRecipients']);
      var from = handleMailbox(item['t:From']);

      var email = {
        id: itemId.id + '|' + itemId.changeKey,
        subject: item['t:Subject'],
        bodyType: item['t:Body']['@']['t:BodyType'],
        body: item['t:Body']['#'],
        size: item['t:Size'],
        dateTimeSent: item['t:DateTimeSent'],
        dateTimeCreated: item['t:DateTimeCreated'],
        toRecipients: toRecipients,
        ccRecipients: ccRecipients,
        from: from,
        isRead: (item['t:IsRead'] == 'true') ? true : false,
        meta: {
          itemId: itemId
        }
      };

      callback(null, email);
    });
  });
};


exports.getFolders = function(id, callback) {
  if (typeof(id) == 'function') {
    callback = id;
    id = 'inbox';
  }

  var soapRequest = 
    '<tns:FindFolder xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
        '<tns:FolderShape>' +
          '<t:BaseShape>Default</t:BaseShape>' +
        '</tns:FolderShape>' +
        '<tns:ParentFolderIds>' + 
          '<t:DistinguishedFolderId Id="inbox"></t:DistinguishedFolderId>' + 
        '</tns:ParentFolderIds>' + 
      '</tns:FindFolder>';

  exports.client.FindFolder(soapRequest, function(err, result) {
    if (err) {
      callback(err)
    }
    
    if (result.ResponseMessages.FindFolderResponseMessage.ResponseCode == 'NoError') {
      var rootFolder = result.ResponseMessages.FindFolderResponseMessage.RootFolder;
    
      callback(null, {});
    }
  });
};


exports.getInboxFolders = function(callback) {

    var  id = 'inbox'
    var soapRequest =
        '<tns:FindFolder Traversal="Shallow" xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
        '<tns:FolderShape>' +
        '<t:BaseShape>AllProperties</t:BaseShape>' +
        '</tns:FolderShape>' +
        '<tns:IndexedPageFolderView Offset="0" BasePoint="Beginning"/>'+
        '<tns:ParentFolderIds>' +
        '<t:DistinguishedFolderId Id="inbox"></t:DistinguishedFolderId>' +
        '</tns:ParentFolderIds>' +
        '</tns:FindFolder>';

    exports.client.FindFolder(soapRequest, function(err, result) {
        if (err) {
            callback(err);
        }
        if(!result || !result.ResponseMessages){
            callback(null, false);
        }else{
            if (result.ResponseMessages.FindFolderResponseMessage.ResponseCode == 'NoError') {
                var rootFolder = result.ResponseMessages.FindFolderResponseMessage.RootFolder;
                
                callback(null, true);
            }else{
                callback(null, false);
            }
        }
    });
};


exports.getAllFolders = function(callback) {
    var soapRequest =
        '<tns:SyncFolderHierarchy xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
        '<tns:FolderShape>' +
        '<t:BaseShape>Default</t:BaseShape>' +
        '</tns:FolderShape>' +
        '</tns:SyncFolderHierarchy>';

    exports.client.SyncFolderHierarchy(soapRequest, function(err, result) {
        if (err) {
            callback(err);
            return;
        }

        if (result.ResponseMessages.SyncFolderHierarchyResponseMessage.ResponseCode == 'NoError') {
            var folders = result.ResponseMessages.SyncFolderHierarchyResponseMessage.Changes;

            callback(null,folders.Create);
        }
    });
};


exports.getEmailsFromFolder = function(start, limit, folderID, sort ,callback) {
    var sortOrder = "Descending";
    if(sort){
        sortOrder = "Ascending";
    }
    var soapRequest =
        '<tns:FindItem Traversal="Shallow">' +
        '<tns:ItemShape>' +
        '<t:BaseShape>Default</t:BaseShape>' +
        '</tns:ItemShape>' +
        '<tns:IndexedPageItemView MaxEntriesReturned="' + limit + '" Offset="' + start + '" BasePoint="Beginning"/>'+
        '<tns:SortOrder>' +
        '<t:FieldOrder Order="' + sortOrder + '">' +
        '<t:FieldURI FieldURI="item:DateTimeReceived"/>' +
        '</t:FieldOrder>' +
        '</tns:SortOrder>' +
        '<tns:ParentFolderIds>' +
        '<t:FolderId Id="' + folderID + '"/>' +
        '</tns:ParentFolderIds>' +
        '</tns:FindItem>';

    exports.client.FindItem(soapRequest, function(err, result) {
        if (err) {
            callback(err);
        }

        if (result.ResponseMessages.FindItemResponseMessage.ResponseCode == 'NoError') {
            var emails = result.ResponseMessages.FindItemResponseMessage.RootFolder.Items.Message;

            callback(null,emails);
        }
    });
};

exports.sendMailWithAttachment = function(subject, body, recipients, files, callback){

  exports.createDraft(subject, body, recipients, function(err, result) {

    var ItemId = result.ResponseMessages.CreateItemResponseMessage.Items.Message.ItemId.attributes.Id;

    exports.createAttachment(ItemId, files, function(err, result1) {

      var ChangeKey = '';

      if (files.length > 1){
        ChangeKey = result1.ResponseMessages.CreateAttachmentResponseMessage[files.length - 1].Attachments.FileAttachment.AttachmentId.attributes.RootItemChangeKey;
      } else {
        ChangeKey = result1.ResponseMessages.CreateAttachmentResponseMessage.Attachments.FileAttachment.AttachmentId.attributes.RootItemChangeKey;
      }

      exports.sendDraft(ItemId, ChangeKey, function(err, result2) {
        callback(err, result2);
      });
    });
  });
};

exports.sendDraft = function(itemId, changeKey, callback) {
  var soapRequest = [
    '<SendItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" SaveItemToFolder="true">',
      '<ItemIds>',
        '<t:ItemId Id="' + itemId + '" ChangeKey="' + changeKey + '"/>',
     ' </ItemIds>',
    '</SendItem>'
    ].join(' '); 

    exports.client.SendItem(soapRequest, function(err, result) {
        if (err) {
            callback(err, null);
        }

        callback(null, result);
    });
};

exports.createAttachment = function(itemId, files, callback) {

  var soapRequest = [
      '<tns:CreateAttachment>',
       '<tns:ParentItemId Id="' + itemId + '" />',
       '<tns:Attachments>'
  ];

  for (var i = 0; i < files.length; i++) {
    soapRequest.push('<t:FileAttachment>');
    soapRequest.push('<t:Name>' + files[i].name + '</t:Name>');
    soapRequest.push('<t:Content>' + files[i].content + '</t:Content>');

    if (files[i].isInline) {
      soapRequest.push('<t:IsInline>true</t:IsInline>');
      soapRequest.push('<t:ContentId>' + files[i].name + '</t:ContentId>');
    }
    soapRequest.push('</t:FileAttachment>');
  }
          
  soapRequest.push('</tns:Attachments>'); 
  soapRequest.push('</tns:CreateAttachment>');
  soapRequest = soapRequest.join(' ');

  exports.client.CreateAttachment(soapRequest, function(err, result) {
      if (err) {
          callback(err, null);
      }

      callback(null, result);
  });
};

exports.createDraft = function( subject, body, recipients, callback) {    
    var soapRequest = [
      '<tns:CreateItem MessageDisposition="SaveOnly">',
          '<tns:Items>',
              '<t:Message>',
                  '<t:ItemClass>IPM.Note</t:ItemClass>',
                  '<t:Subject>' + subject + '</t:Subject>',
                  '<t:Body BodyType="HTML">' + body + '</t:Body>',
                  '<t:ToRecipients>'
    ];

    for (var i=0; i < recipients.length; i++) {
      soapRequest.push('<t:Mailbox>');
      soapRequest.push('<t:EmailAddress>' + recipients[i] + '</t:EmailAddress>');
      soapRequest.push('</t:Mailbox>');
    }
          
    soapRequest.push('</t:ToRecipients>'); 
    soapRequest.push('</t:Message>'); 
    soapRequest.push('</tns:Items>'); 
    soapRequest.push('</tns:CreateItem>');
    soapRequest = soapRequest.join(' ');

    exports.client.CreateItem(soapRequest, function(err, result) {
        if (err) {
            callback(err, null);
        }

        callback(null, result);

    });

};

exports.sendMail = function(subject, body, emailTo, emailCc, emailBcc, nameTo, emailFrom, nameFrom ,callback) {

    var soapRequest = [
      '<tns:CreateItem MessageDisposition="SendAndSaveCopy">',
        '<tns:Items>',
        
          '<t:Message>',
          '<t:ItemClass>IPM.Note</t:ItemClass>',
            '<t:Subject>' + subject + '</t:Subject>',
            '<t:Body BodyType="HTML">' + body + '</t:Body>',
            '<t:ToRecipients>'
    ];
            for (var i=0; i < emailTo.length; i++) {
              soapRequest.push('<t:Mailbox>');
              soapRequest.push('<t:EmailAddress>' + emailTo[i] + '</t:EmailAddress>');
              soapRequest.push('</t:Mailbox>');
            }
            soapRequest.push('</t:ToRecipients>');
            soapRequest.push('<t:CcRecipients>');
             for (var j=0; j < emailCc.length; j++) {
              soapRequest.push('<t:Mailbox>');
              soapRequest.push('<t:EmailAddress>' + emailCc[j] + '</t:EmailAddress>');
              soapRequest.push('</t:Mailbox>');
            }
            soapRequest.push('</t:CcRecipients>');
            soapRequest.push('<t:BccRecipients>');
            for (var k=0; k < emailBcc.length; k++) {
              soapRequest.push('<t:Mailbox>');
              soapRequest.push('<t:EmailAddress>' + emailBcc[k] + '</t:EmailAddress>');
              soapRequest.push('</t:Mailbox>');
            }
            soapRequest.push('</t:BccRecipients>');
            soapRequest.push('<t:From>');
            soapRequest.push(  '<t:Mailbox>');
            soapRequest.push(   '<t:Name>' + nameFrom + '</t:Name>');
            soapRequest.push(   '<t:EmailAddress>' + emailFrom + '</t:EmailAddress>');
            soapRequest.push(  '</t:Mailbox>');
            soapRequest.push('</t:From>');
          soapRequest.push('</t:Message>');
        soapRequest.push('</tns:Items>');
      soapRequest.push('</tns:CreateItem>');
      
      soapRequest = soapRequest.join(' ');
      
    exports.client.CreateItem(soapRequest, function(err, result) {
        if (err) {
            callback(err, null);
        }

        if (result.ResponseMessages) {
          if (result.ResponseMessages.CreateItemResponseMessage.ResponseCode == 'NoError') {
              var emails = result.ResponseMessages.CreateItemResponseMessage.ResponseCode;

              callback(null,emails);
          }
        } else {
          if (result.ResponseMessage.CreateItemResponseMessage.ResponseCode == 'NoError') {
              var emails = result.ResponseMessage.CreateItemResponseMessage.ResponseCode;

              callback(null,emails);
          }
        }
    });
};