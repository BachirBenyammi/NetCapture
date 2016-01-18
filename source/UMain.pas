unit UMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, OleCtrls, SHDocVw, ExtCtrls, MSHTML_TLB, ActiveX, IniFiles,
  ComCtrls, FileCtrl, ShellApi, Menus, Buttons, ToolWin, ActnList, ImgList,
  Gauges, CommCtrl;

Const
  WM_MYMESSAGE=WM_USER+100;

type
  TMainForm = class(TForm)
    Timer: TTimer;
    sBar: TStatusBar;
    PMenu: TPopupMenu;
    PMShow: TMenuItem;
    PMAbout: TMenuItem;
    PMLine: TMenuItem;
    PMExit: TMenuItem;
    PMActivate: TMenuItem;
    OpenDialog1: TOpenDialog;
    ActionList1: TActionList;
    actBack: TAction;
    actNext: TAction;
    actStop: TAction;
    actRefresh: TAction;
    actHome: TAction;
    actSearch: TAction;
    actPrint: TAction;
    actSave: TAction;
    actOpen: TAction;
    actHTML: TAction;
    actText: TAction;
    actLinks: TAction;
    actIE: TAction;
    actOK: TAction;
    PnlTop: TPanel;
    cbStayOnTop: TCheckBox;
    btnExit: TButton;
    btnAbout: TButton;
    btnHide: TButton;
    btnOpen: TButton;
    btnSelect: TButton;
    btnActivate: TButton;
    actFull: TAction;
    Splitter1: TSplitter;
    PnlRight: TPanel;
    WebBrowser1: TWebBrowser;
    ToolButton9: TToolButton;
    ToolButton8: TToolButton;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton14: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    ToolButton23: TToolButton;
    PnlLeft: TPanel;
    lbUrls: TListBox;
    ToolButton20: TToolButton;
    ToolButton19: TToolButton;
    ToolButton17: TToolButton;
    Panel1: TPanel;
    ToolBar1: TToolBar;
    ImageList2: TImageList;
    Gauge: TGauge;
    ToolButton10: TToolButton;
    ToolButton15: TToolButton;
    ToolButton18: TToolButton;
    ToolBar2: TToolBar;
    ComboBoxURL: TComboBox;
    ToolButton16: TToolButton;
    procedure TimerTimer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnActivateClick(Sender: TObject);
    procedure btnSelectClick(Sender: TObject);
    procedure btnOpenClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure btnExitClick(Sender: TObject);
    procedure btnAboutClick(Sender: TObject);
    procedure btnHideClick(Sender: TObject);
    procedure lbUrlsClick(Sender: TObject);
    procedure cbStayOnTopClick(Sender: TObject);

    procedure WebBrowser1CommandStateChange(Sender: TObject;
      Command: Integer; Enable: WordBool);
    procedure actBackExecute(Sender: TObject);
    procedure actNextExecute(Sender: TObject);
    procedure actStopExecute(Sender: TObject);
    procedure actRefreshExecute(Sender: TObject);
    procedure actHomeExecute(Sender: TObject);
    procedure actOpenExecute(Sender: TObject);
    procedure actLinksExecute(Sender: TObject);
    procedure actOKExecute(Sender: TObject);
    procedure actSearchExecute(Sender: TObject);
    procedure actPrintExecute(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actIEExecute(Sender: TObject);
    procedure actHTMLExecute(Sender: TObject);
    procedure actTextExecute(Sender: TObject);
    procedure WebBrowser1TitleChange(Sender: TObject;
      const Text: WideString);
    procedure WebBrowser1StatusTextChange(Sender: TObject;
      const Text: WideString);
    procedure WebBrowser1ProgressChange(Sender: TObject; Progress,
      ProgressMax: Integer);
    procedure ComboBoxURLKeyPress(Sender: TObject; var Key: Char);
    procedure ComboBoxURLDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure ComboBoxURLDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure actFullExecute(Sender: TObject);
    procedure Splitter1Moved(Sender: TObject);
    procedure FormShow(Sender: TObject);
  Public
    PageStr, UrlList, FilesList : TStringList;
    procedure Capture;
    procedure GetPath;
    procedure SetPath(NewPath : string);
    function CheckSlash(Path : string):string;
    procedure InitialSysTray;
    procedure ShowHide;
    procedure TrayMessage(var Msg: TMessage); message WM_MYMESSAGE;
    procedure WMHotkey(Var msg:TWMHotkey); message WM_HOTKEY;
    function GetHTMLCode(WB: IWebbrowser2; ACode: TStrings): Boolean;
 end;

var
  MainForm: TMainForm;
  UrlCap : integer = 0;
  IEInst : integer = 0;
  ExePath, Path, Dest, Jour, Heure: string;
  PNotifyStruc: PNotifyIconData;
  WebB:WebBrowser;

implementation

uses Unit2, Unit3;

{$R *.dfm}

procedure TMainForm.InitialSysTray;
var
  hinte: string;
  j: integer;
begin
  Icon := Application.Icon;
  hinte := Application.Title;
  New(PNotifyStruc);
  PNotifyStruc^.cbSize := SizeOf(PNotifyStruc^);
  PNotifyStruc^.Wnd := Handle;
  PNotifyStruc^.uID := 1;
  PNotifyStruc^.uFlags := NIF_ICON or NIF_TIP or NIF_MESSAGE;
  PNotifyStruc^.uCallbackMessage := WM_MYMESSAGE;
  PNotifyStruc^.hIcon := Icon.Handle;
  for j := 0 to length(hinte) - 1 do
    PNotifyStruc^.szTip[j] := hinte[j + 1];
  PNotifyStruc^.szTip[length(hinte)] := #0;
  Shell_NotifyIcon(NIM_ADD,PNotifyStruc);
end;

procedure TMainForm.TrayMessage(var Msg: TMessage);
var
  Mouse :TPoint;
begin
  case Msg.LParam of
    WM_RBUTTONDOWN:
      begin
        GetCursorPos(Mouse);
        SetForegroundWindow(Handle);
        PMenu.Popup(Mouse.x,Mouse.y);
      end;
    WM_LBUTTONDOWN:
      ShowHide
  end;
end;

procedure TMainForm.WMHotkey(Var msg:TWMHotkey);
Begin
  ShowHide
end;

procedure TmainForm.ShowHide;
begin
  case Visible of
    true:
      begin
        Visible := false;
        PMShow.Enabled := true;
      end;
    false:
      begin
        Visible := true;
        PMShow.Enabled := false;
        ShowWindow(Application.Handle, SW_Hide);
      end;
  end;
end;

function TMainForm.CheckSlash(Path : string):string;
var
 NewPath : string;
begin
  NewPath := Path;
  if NewPath[length(NewPath)] <> '\' then
    NewPath := NewPath + '\';
  result := NewPath;
end;

procedure TMainForm.GetPath;
var
  iniFile : TIniFile;
  NewPath : string;
begin
  iniFile := TIniFile.Create(ExePath + 'config.ini');
  NewPath := iniFile.ReadString('Options', 'Dest', Path);
  Path := CheckSlash(NewPath);
  sBar.Panels[5].Text := Path;
  iniFile.Free;
end;

procedure TMainForm.SetPath(NewPath : string);
var
  iniFile : TIniFile;
begin
 iniFile := TIniFile.Create(ExePath + 'config.ini');
 iniFile.WriteString('Options', 'Dest', NewPath);
 Path := CheckSlash(NewPath);
 sBar.Panels[5].Text := Path;
 iniFile.Free;
end;

function TMainForm.GetHTMLCode(WB: IWebbrowser2; ACode: TStrings): Boolean;
var
  ps: IPersistStreamInit;
  s: string;
  ss: TStringStream;
  sa: IStream;
begin
  ps := WB.document as IPersistStreamInit;
  s := '';
  ss := TStringStream.Create(s);
  try
    sa:= TStreamAdapter.Create(ss, soReference) as IStream;
    Result := Succeeded(ps.Save(sa, Bool(True)));
    if Result then ACode.Add(ss.Datastring);
  finally
    ss.Free;
  end;
end;

procedure TMainForm.Capture;
var
  ShellWindow: IShellWindows;
  WB: IWebbrowser2;
  spDisp: IDispatch;
  IDoc1: IHTMLDocument2;
  k: Integer;
begin
  ShellWindow := CoShellWindows.Create;
  IEInst := ShellWindow.Count -1;
  for k := 0 to IEInst do
    begin
      spDisp := ShellWindow.Item(k);
      if spDisp = nil then Continue;
      spDisp.QueryInterface(iWebBrowser2, WB);
      if WB <> nil then
        begin
          Jour := FormatDateTime('d-mm-yy',Now);
          Dest := Path + Jour + '\';
          if not DirectoryExists(Dest) then
            MkDir(Dest);
          WB.Document.QueryInterface(IHTMLDocument2, iDoc1);
          sBar.Panels[4].Text := 'IE Inst.: ' + IntToStr(IEInst);
          if iDoc1 <> nil then
            begin
              WB := ShellWindow.Item(k) as IWebbrowser2;
              if UrlList.IndexOf(WB.LocationURL) <> -1  then Continue;
              if WB.ReadyState <> READYSTATE_COMPLETE then Continue;
              Heure := FormatDateTime('hh-mm-ss',Now);
              PageStr.Clear;
              PageStr.Add('Title: ' + WB.LocationName + '<br>');
              PageStr.Add('URL: ' + WB.LocationURL + '<br>');
              PageStr.Add('Time:' + Jour + ', ' + Heure + '<br>');
              GetHTMLCode(WB, PageStr);
              UrlList.Add(WB.LocationURL);
              lbUrls.Items.Add(WB.LocationURL);
              FilesList.Add(Dest + Heure + '.htm');
              Inc(UrlCap);
              sBar.Panels[3].Text := 'Captured: ' + IntToStr(UrlCap);
              PageStr.SaveToFile(Dest + Heure + '.htm');
            end;
        end;
    end;
end;

procedure TMainForm.TimerTimer(Sender: TObject);
begin
  Timer.Enabled := false;
  Capture;
  Timer.Enabled := true;
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  ExePath := CheckSlash(ExtractFilePath(Application.ExeName));
  Path := ExePath;
  GetPath;

  PageStr := TStringList.Create;
  UrlList := TStringList.Create;
  FilesList := TStringList.Create;

  InitialSysTray;
  PMShow.Caption := Caption;
  RegisterHotkey(Handle, 1, Mod_Control + Mod_Alt, 38);
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  Timer.Enabled := False;
  PageStr.free;
  UrlList.Free;
  FilesList.Free;
  
  UnRegisterHotkey(Handle, 1);
  if (PNotifyStruc <> NIL) then
    begin
      Shell_NotifyIcon(NIM_DELETE,PNotifyStruc);
      Dispose(PNotifyStruc);
    end;
end;

procedure TMainForm.btnActivateClick(Sender: TObject);
begin
  with Timer do
    begin
      Enabled := not Enabled;
      if Enabled then
        begin
          btnActivate.Caption := '&Disabled';
          PMActivate.Caption := '&Disabled';
          Caption := 'NetCapture [Enabled] - BENBAC SOFT';
          Icon.LoadFromFile(Path + 'red.ico');
          PNotifyStruc^.hIcon := Icon.Handle;
          Shell_NotifyIcon(NIM_Modify,PNotifyStruc);
        end
      else
        begin
          btnActivate.Caption := '&Enabled';
          PMActivate.Caption := '&Enabled';
          Caption := 'NetCapture [Disabled] - BENBAC SOFT';
          Icon := Application.Icon;
          PNotifyStruc^.hIcon := Icon.Handle;
          Shell_NotifyIcon(NIM_MODIFY, PNotifyStruc);
        end;
    end;
end;

procedure TMainForm.btnSelectClick(Sender: TObject);
var
  Dir : string;
begin
  Dir := Path;
  if SelectDirectory('Select a folder ', '', Dir) then
    SetPath(Dir);
end;

procedure TMainForm.btnOpenClick(Sender: TObject);
begin
  ShellExecute(0, 'open', pchar(Dest), nil, nil, sw_show);
end;

procedure TMainForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  ShowHide;
  CanClose := false;
end;

procedure TMainForm.btnExitClick(Sender: TObject);
begin
  if MessageDlg('Do you want to exit ' + Application.Title + ' ?', mtConfirmation,
    [mbYes, mbNo], 0) = mrYes then
    Application.Terminate;
end;

procedure TMainForm.btnAboutClick(Sender: TObject);
begin
  MessageDlg(Application.Title, mtInformation, [mbOk], 0)
end;

procedure TMainForm.btnHideClick(Sender: TObject);
begin
  ShowHide
end;

procedure TMainForm.lbUrlsClick(Sender: TObject);
begin
  if lbUrls.ItemIndex <> -1 then
  //ShellExecute(0, 'open', pchar(FilesList.Strings[lbUrls.ItemIndex]), nil, nil, sw_show);
  try
    WebBrowser1.Navigate(FilesList.Strings[lbUrls.ItemIndex],
      EmptyParam, EmptyParam, EmptyParam, EmptyParam);
  except
  end;
end;

procedure TMainForm.cbStayOnTopClick(Sender: TObject);
begin
  if cbStayOnTop.Checked then
    FormStyle := fsStayOnTop
  else
    FormStyle := fsNormal;

end;

procedure TMainForm.WebBrowser1CommandStateChange(Sender: TObject;
  Command: Integer; Enable: WordBool);
begin
  case Command of
    CSC_NAVIGATEFORWARD : actNext.Enabled := Enable;
    CSC_NAVIGATEBACK : actBack.Enabled:= Enable;
  end;
end;

procedure TMainForm.actBackExecute(Sender: TObject);
begin
  try WebBrowser1.GoBack except end;
end;

procedure TMainForm.actNextExecute(Sender: TObject);
begin
  try WebBrowser1.GoForward except end;
end;

procedure TMainForm.actStopExecute(Sender: TObject);
begin
 try WebBrowser1.Stop except end;
end;

procedure TMainForm.actRefreshExecute(Sender: TObject);
begin
  try WebBrowser1.Refresh except end;
end;

procedure TMainForm.actHomeExecute(Sender: TObject);
begin
  try WebBrowser1.GoHome except end;
end;

procedure TMainForm.actOpenExecute(Sender: TObject);
begin
  if Opendialog1.Execute then
  begin
    ComboBoxURL.Text:='file://'+OpenDialog1.FileName;
    actOKExecute(nil)
  end;
end;

procedure TMainForm.actLinksExecute(Sender: TObject);
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

procedure TMainForm.actOKExecute(Sender: TObject);
begin
  try
    WebBrowser1.Navigate(ComboBoxURL.Text,
      EmptyParam, EmptyParam, EmptyParam, EmptyParam);
    ComboBoxURL.Items.Add(ComboBoxURL.Text);
  except
  end;
end;

procedure TMainForm.actSearchExecute(Sender: TObject);
begin
  try WebBrowser1.GoSearch except end;
end;

procedure TMainForm.actPrintExecute(Sender: TObject);
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
      sBar.Panels[2].Text:='impossible d''imprimer';
  except
  end;
end;

procedure TMainForm.actSaveExecute(Sender: TObject);
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
      sBar.Panels[2].Text:='impossible de sauver';
  except
  end;
end;

procedure TMainForm.actIEExecute(Sender: TObject);
begin
  WebB := CoInternetExplorer.Create;
  WebB.visible := true;
  WebB.navigate(ComboBoxURL.text, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
end;

procedure TMainForm.actHTMLExecute(Sender: TObject);
begin
  Form2.show;
  try
    Form2.RichEdit1.Text := WebBrowser1.OleObject.Document.body.innerHTML;
  except
  end;
end;

procedure TMainForm.actTextExecute(Sender: TObject);
begin
  Form3.show;
  try
    Form3.RichEdit1.Text := WebBrowser1.OleObject.Document.body.innerText;
  except end;
end;

procedure TMainForm.WebBrowser1TitleChange(Sender: TObject;
  const Text: WideString);
begin
  Caption := Text + '-' + Application.Title;
  ComboBoxURL.Text := WebBrowser1.LocationURL;
end;

procedure TMainForm.WebBrowser1StatusTextChange(Sender: TObject;
  const Text: WideString);
begin
  sBar.Panels[0].Text := Text;
end;

procedure TMainForm.WebBrowser1ProgressChange(Sender: TObject; Progress,
  ProgressMax: Integer);
begin
  if ProgressMax > 0 then
    begin
      Gauge.MaxValue := ProgressMax;
      Gauge.AddProgress(Progress);
    end
  else
    Gauge.Progress := 0;
end;

procedure TMainForm.ComboBoxURLKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
  begin
    actOKExecute(nil);
    key := #0;
  end;
end;

procedure TMainForm.ComboBoxURLDragDrop(Sender, Source: TObject; X,
  Y: Integer);
begin
  ComboBoxURL.Text := WebBrowser1.StatusText;
{var
  item: Variant;
begin
 if (Webbrowser1.ReadyState and READYSTATE_INTERACTIVE) = 3 then
  begin
    item:= WebBrowser1.OleObject.Document.elementFromPoint(x, y);
    ComboBoxURL.Text := item.Value;
  end;
}
end;

procedure TMainForm.ComboBoxURLDragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := Source = WebBrowser1;
end;

procedure TMainForm.actFullExecute(Sender: TObject);
begin
  if actFull.Checked then
    begin
      actFull.Checked := false;
      actFull.Caption := 'Full Screen';
      actFull.Hint := 'Full Screen';
      PnlTop.Visible := true;
      PnlLeft.Visible := true;
    end
  else
    begin
      actFull.Checked := true;
      actFull.Caption := 'Normal Screen';
      actFull.Hint := 'Normal Screen';
      PnlTop.Visible := false;
      PnlLeft.Visible := false;
      PnlRight.Width := Width;
      PnlRight.Height := Height
    end;
end;

procedure TMainForm.Splitter1Moved(Sender: TObject);
begin
  ComboBoxURL.Width := Width  - PnlRight.Left - ComboBoxURL.Left  - 65 ;
end;

procedure TMainForm.FormShow(Sender: TObject);
var
  rect : TRect;
begin
 sBar.perform(SB_GETRECT, 1, integer(@rect));
 with Gauge do
  begin
    parent := sBar;
    top := rect.top;
    left := rect.left;
    width := rect.right - rect.left;
    height := rect.bottom - rect.top;
    Visible := true;
  end;
end;

end.
