unit UBrowser;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, OleCtrls, ToolWin, ComCtrls, ComObj, SHDocVw, ActnList, ImgList,
  Buttons, ExtCtrls;

type
  TFormBrowser = class(TForm)
    OpenDialog1: TOpenDialog;
    StatusBar1: TStatusBar;
    ImageList1: TImageList;
    ActionList1: TActionList;
    actBack: TAction;
    actNext: TAction;
    actStop: TAction;
    actRefresh: TAction;
    actHome: TAction;
    actSearch: TAction;
    actPrint: TAction;
    actSave: TAction;
    actBrowse: TAction;
    actHTML: TAction;
    actText: TAction;
    actAbout: TAction;
    actLinks: TAction;
    actOLE: TAction;
    actOK: TAction;
    ToolBar2: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    ToolButton10: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    ToolButton14: TToolButton;
    ToolButton15: TToolButton;
    ToolButton16: TToolButton;
    ToolButton17: TToolButton;
    ToolButton18: TToolButton;
    ToolButton19: TToolButton;
    ProgressBar1: TProgressBar;
    ToolBar1: TToolBar;
    Label1: TLabel;
    ComboBoxURL: TComboBox;
    SpeedButton1: TSpeedButton;
    ToolButton20: TToolButton;
    ToolButton21: TToolButton;
    Panel1: TPanel;
    WebBrowser1: TWebBrowser;
    procedure ComboBoxURLKeyPress(Sender: TObject; var Key: Char);
    procedure WebBrowser1CommandStateChange(Sender: TObject;
      Command: Integer; Enable: WordBool);
    procedure actBackExecute(Sender: TObject);
    procedure actNextExecute(Sender: TObject);
    procedure actStopExecute(Sender: TObject);
    procedure actRefreshExecute(Sender: TObject);
    procedure actHomeExecute(Sender: TObject);
    procedure actBrowseExecute(Sender: TObject);
    procedure actLinksExecute(Sender: TObject);
    procedure actOKExecute(Sender: TObject);
    procedure actSearchExecute(Sender: TObject);
    procedure actPrintExecute(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actOLEExecute(Sender: TObject);
    procedure actHTMLExecute(Sender: TObject);
    procedure actTextExecute(Sender: TObject);
    procedure actAboutExecute(Sender: TObject);
    procedure WebBrowser1TitleChange(Sender: TObject;
      const Text: WideString);
    procedure WebBrowser1StatusTextChange(Sender: TObject;
      const Text: WideString);
    procedure WebBrowser1ProgressChange(Sender: TObject; Progress,
      ProgressMax: Integer);
  end;

var
  FormBrowser: TFormBrowser;
  WebB:WebBrowser;

implementation

uses Unit2, Unit3, UMain;

{$R *.DFM}

procedure TFormBrowser.ComboBoxURLKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
  begin
    actOKExecute(nil);
    key := #0;
  end;
end;

procedure TFormBrowser.WebBrowser1CommandStateChange(Sender: TObject;
  Command: Integer; Enable: WordBool);
begin
  case Command of
    CSC_NAVIGATEFORWARD : actNext.Enabled := Enable;
    CSC_NAVIGATEBACK : actBack.Enabled:= Enable;
  end;
end;

procedure TFormBrowser.actBackExecute(Sender: TObject);
begin
  try WebBrowser1.GoBack except end;
end;

procedure TFormBrowser.actNextExecute(Sender: TObject);
begin
  try WebBrowser1.GoForward except end;
end;

procedure TFormBrowser.actStopExecute(Sender: TObject);
begin
 try WebBrowser1.Stop except end;
end;

procedure TFormBrowser.actRefreshExecute(Sender: TObject);
begin
  try WebBrowser1.Refresh except end;
end;

procedure TFormBrowser.actHomeExecute(Sender: TObject);
begin
  try WebBrowser1.GoHome except end;
end;

procedure TFormBrowser.actBrowseExecute(Sender: TObject);
begin
  if Opendialog1.Execute then
  begin
    ComboBoxURL.Text:='file://'+OpenDialog1.FileName;
    actOKExecute(nil)
  end;
end;

procedure TFormBrowser.actLinksExecute(Sender: TObject);
var
  i : Integer;
begin
  Form2.show;
  try
    For i:=0 to Webbrowser1.OleObject.Document.links.length-1 Do
      Form2.RichEdit1.lines.add(Webbrowser1.OleObject.Document.links.item(i));
  Except
  end;
end;

procedure TFormBrowser.actOKExecute(Sender: TObject);
begin
  try
    WebBrowser1.Navigate(ComboBoxURL.Text,
      EmptyParam, EmptyParam, EmptyParam, EmptyParam);
  except
  end;
end;

procedure TFormBrowser.actSearchExecute(Sender: TObject);
begin
  try WebBrowser1.GoSearch except end;
end;

procedure TFormBrowser.actPrintExecute(Sender: TObject);
var
  eQuery:variant;
  //pcmdf: OLECMDF;  
begin
  try
    eQuery:=WebBrowser1.QueryStatusWB(OLECMDID_PRINT);
  //eQuery:=WebBrowser1.QueryStatusWB(OLECMDID_PRINT,pcmdf);
    if eQuery and OLECMDF_ENABLED then
       WebBrowser1.ExecWB(OLECMDID_PRINT,
        OLECMDEXECOPT_PROMPTUSER, EmptyParam, EmptyParam)
    else
      StatusBar1.Panels[2].Text:='impossible d''imprimer';
  except
  end;
end;

procedure TFormBrowser.actSaveExecute(Sender: TObject);
var
  eQuery: variant;
  //pcmdf: OLECMDF;
begin
  try
    eQuery:=WebBrowser1.QueryStatusWB(OLECMDID_SAVEAS);
  //eQuery:=WebBrowser1.QueryStatusWB(OLECMDID_SAVEAS,pcmdf);
    If eQuery and OLECMDF_ENABLED then
       WebBrowser1.ExecWB(OLECMDID_SAVEAS,
        OLECMDEXECOPT_PROMPTUSER,EmptyParam,EmptyParam)
    else
      StatusBar1.Panels[2].Text:='impossible de sauver';
  except
  end;
end;

procedure TFormBrowser.actOLEExecute(Sender: TObject);
begin
  WebB := CoInternetExplorer.Create;
  WebB.visible := true;
  WebB.navigate(ComboBoxURL.text, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
end;

procedure TFormBrowser.actHTMLExecute(Sender: TObject);
begin
  Form2.show;
  try
    Form2.RichEdit1.Text := WebBrowser1.OleObject.Document.body.innerHTML;
  except
  end;
end;

procedure TFormBrowser.actTextExecute(Sender: TObject);
begin
  Form3.show;
  try
    Form3.RichEdit1.Text := WebBrowser1.OleObject.Document.body.innerText;
  except end;
end;

procedure TFormBrowser.actAboutExecute(Sender: TObject);
var
  CodeHTML : string;
begin
  CodeHTML := '<center><B><H>' + Application.Title +'</H></B></center>';
  WebBrowser1.Navigate('about:' + CodeHTML, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
end;

procedure TFormBrowser.WebBrowser1TitleChange(Sender: TObject;
  const Text: WideString);
begin
  Caption := Text + '-' + Application.Title;
  ComboBoxURL.Text := WebBrowser1.LocationURL;
end;

procedure TFormBrowser.WebBrowser1StatusTextChange(Sender: TObject;
  const Text: WideString);
begin
  StatusBar1.Panels[0].Text := Text;
end;

procedure TFormBrowser.WebBrowser1ProgressChange(Sender: TObject; Progress,
  ProgressMax: Integer);
begin
  ProgressBar1.Max := ProgressMax;
  ProgressBar1.Position := Progress;
  if ProgressMax > 0 then
    StatusBar1.Panels[1].Text := 'Loaded: ' +
      IntToStr(Round (Progress * 100 / ProgressMax)) + '%';
end;

end.
