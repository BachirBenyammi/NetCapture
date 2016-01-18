unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, SynCompletionProposal, SynEditHighlighter, SynHighlighterHtml,
  SynEdit, SynMemo;

type
  TFormViewer = class(TForm)
    SynHTML: TSynHTMLSyn;
    SynEditor: TSynEdit;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormViewer: TFormViewer;

implementation

{$R *.dfm}

end.
