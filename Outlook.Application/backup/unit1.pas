unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, windows,comobj;

type

  { TForm1 }

  TForm1 = class(TForm)
    procedure FormCreate(Sender: TObject);
  private

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

  procedure TForm1.FormCreate(Sender: TObject);
  const olMailItem = $00000000;
  var
    OutlookApp , oNameSpace, Inbox, MailItem : OleVariant;
    Mail: Variant;
    i: Integer;
  begin
    try

    try
         OutlookApp   := GetActiveOleObject('Outlook.Application');
    except
      try
           OutlookApp   := CreateOleObject('Outlook.Application');

      except
        ShowMessage('Outlook Application not installed !!');
      end;
    end;

    //oNameSpace := Outlook.GetNamespace('MAPI');
    //oNameSpace.Logon('', '', False, False);   // not sure if this is necessary
    //Inbox := oNameSpace.GetDefaultFolder(6);
    //for i := 1 to Inbox.Folders.Count do
    //  ShowMessage(Inbox.Folders[i].Name);


    //Mail := OutlookApp .CreateItem(olMailItem);
    //Mail.To := 'receiver1@xyz.com' + ';' + 'receiver2@xyz.com';
    //Mail.Subject := 'your subject';
    //Mail.Display; //Mail.Send; if you want to send directly




     try
      //wsRecep:=UTF8Decode(mail_id);
      MailItem := OutlookApp.CreateItem(olMailItem);
      //MailItem.Recipients.Add(wsRecep);
      MailItem.Recipients.Add( 'someone@somewhere.com' );
      MailItem.Subject := 'Training Enrollemnt Details';
      MailItem.Body := ('You have enrolled to the training courses');
      MailItem.Display;
      //ShowMessage('Mail sent Successfully');
    finally
      OutlookApp    := VarNull;
    end;










    Except
    on E : Exception do
      ShowMessage(E.ClassName+' error raised, with message : '+E.Message);end;

  end;

end.

