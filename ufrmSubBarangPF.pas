unit ufrmSubBarangPF;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, AdvPanel, ComCtrls, StdCtrls, AdvEdit,SqlExpr, Menus,
  cxLookAndFeelPainters, cxButtons,StrUtils, cxGraphics, cxLookAndFeels,
  dxSkinsCore, dxSkinsDefaultPainters, cxControls, cxContainer, cxEdit,
  cxTextEdit, cxMaskEdit, cxDropDownEdit, cxLookupEdit, cxDBLookupEdit,
  cxDBExtLookupComboBox,DBClient, MyAccess;

type
  TfrmSubBarangPF = class(TForm)
    AdvPanel1: TAdvPanel;
    Label2: TLabel;
    edtKode: TAdvEdit;
    Label3: TLabel;
    edtNama: TAdvEdit;
    AdvPanel3: TAdvPanel;
    cxButton8: TcxButton;
    AdvPanel2: TAdvPanel;
    lbljudul: TLabel;
    cxButton2: TcxButton;
    cxButton1: TcxButton;
    Label1: TLabel;
    cxLookupSubKode: TcxExtLookupComboBox;
    Label4: TLabel;
    edtHNA: TAdvEdit;
    procedure refreshdata;
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure loaddata(akode:string) ;
    procedure simpandata;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    function getmaxkode:string;
    procedure cxButton1Click(Sender: TObject);
    procedure cxButton8Click(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure edtKodeExit(Sender: TObject);
    procedure initview;
    

  private
    FCDSSubBarang: TClientDataset;
    FFLAGEDIT: Boolean;
    FID: string;
    FTingkat : Integer;
    function GetCDSSubBarang: TClientDataset;

    { Private declarations }
  public
    property CDSSubBarang: TClientDataset read GetCDSSubBarang write FCDSSubBarang;
    property FLAGEDIT: Boolean read FFLAGEDIT write FFLAGEDIT;
    property ID: string read FID write FID;
    property Tingkat: Integer read FTingkat write FTingkat;
    { Public declarations }
  end;

var
  frmSubBarangPF: TfrmSubBarangPF;

implementation
uses MAIN,uModuleConnection,uFrmbantuan,Ulib, DB;

{$R *.dfm}

procedure TfrmSubBarangPF.refreshdata;
begin
  FID:='';
  flagedit := false;
  cxLookupSubKode.Enabled := true;
  edtNama.Clear;
  edtHNA.Text := '0';
  edtNama.SetFocus;
end;
procedure TfrmSubBarangPF.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

  if Key = VK_F8 then
  begin
      Release;
  end;


  if Key= VK_F10 then
  begin
    try
      if (FLAGEDIT) and ( not cekedit(frmMenu.KDUSER,self.name)) then
        begin
           MessageDlg('Anda tidak berhak Edit di Modul ini',mtWarning, [mbOK],0);
           Exit;
        End;
         if (not FLAGEDIT) and ( not cekinsert(frmMenu.KDUSER,self.name)) then
        begin
           MessageDlg('Anda tidak berhak Insert di Modul ini',mtWarning, [mbOK],0);;
           Exit;
        End;
      
      if MessageDlg('Yakin ingin simpan ?',mtCustom,
                                  [mbYes,mbNo], 0)= mrNo
      then Exit ;

      simpandata;
      refreshdata;
   except
     ShowMessage('Gagal Simpan');
     // xRollback(frmMenu.conn);
     Exit;
   end;
    // xCommit(frmMenu.conn);
  end;
end;

procedure TfrmSubBarangPF.FormKeyPress(Sender: TObject; var Key: Char);
begin
   if Key = #13 then
      SelectNext(ActiveControl,True,True);
end;

procedure TfrmSubBarangPF.loaddata(akode:string) ;
var
  s: string;
  tsql : TmyQuery;
begin
  s:= 'select kode_grouppf,nama_grouppf,kode_deptpf,HNA_grouppf from tgrouppf where kode_grouppf = ' + Quot(akode) ;
tsql := xOpenQuery(s,frmMenu.conn);
with tsql do
begin
  try
    if not Eof then
    begin
      FLAGEDIT := True;
      edtKode.Enabled := False;
      edtNama.Text := fieldbyname('nama_grouppf').AsString;
      cxLookupSubKode.EditValue := fieldbyname('kode_deptpf').AsString;
      FID :=fieldbyname('kode_grouppf').Asstring;
      edtHNA.Text :=fieldbyname('hna_grouppf').AsString;
    end
    else
     FLAGEDIT := False;

  finally
    Free;
  end;
end;

end;


procedure TfrmSubBarangPF.simpandata;
var
  s:string;
begin
if FLAGEDIT then
begin

  s:='update tgrouppf set '
    + ' nama_grouppf = ' + Quot(edtNama.Text)+','
    + ' kode_deptpf='+ IntToStr(CDSSubBarang.Fields[1].AsInteger)+','
    + ' hna_grouppf='+ StringReplace(edtHna.Text,',','',[rfReplaceAll])
    + ' where kode_grouppf = ' + quot(FID) + ';';
   // xExecQuery(s,frmMenu.conn);
EnsureConnected(frmMenu.conn);
ExecSQLDirect(frmMenu.conn, s);
end
else
begin

  s :=  ' insert into tgrouppf '
             + ' (kode_grouppf,nama_grouppf,kode_deptpf,hna_grouppf) '
             + ' values ( '
             + Quot(edtKode.Text) + ','
             + Quot(edtNama.Text) + ' ,'
             + IntToStr(CDSSubBarang.Fields[1].AsInteger)+','
             + StringReplace(edtHNA.Text,',','',[rfReplaceAll])
             + ');';
    // xExecQuery(s,frmMenu.conn);
EnsureConnected(frmMenu.conn);
ExecSQLDirect(frmMenu.conn, s);
    edtkode.Text := getmaxkode;
end;


end;


procedure TfrmSubBarangPF.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   Action := caFree;
   Release;
end;

function TfrmSubBarangPF.getmaxkode:string;
var
  s:string;
begin

   s:= 'select  max(kode_grouppf) from tgrouppf ';
   with xOpenQuery(s,frmMenu.conn) do
  begin
    try
      if Fields[0].AsString = '' then
         result:= '1'
      else
         result:= vartostr(IntToStr(fields[0].AsInteger+1));

    finally
      free;
    end;
  end;

end;

procedure TfrmSubBarangPF.cxButton1Click(Sender: TObject);
begin
    try
      if (FLAGEDIT) and ( not cekedit(frmMenu.KDUSER,self.name)) then
        begin
           MessageDlg('Anda tidak berhak Edit di Modul ini',mtWarning, [mbOK],0);
           Exit;
        End;
         if (not FLAGEDIT) and ( not cekinsert(frmMenu.KDUSER,self.name)) then
        begin
           MessageDlg('Anda tidak berhak Insert di Modul ini',mtWarning, [mbOK],0);;
           Exit;
        End;

      if MessageDlg('Yakin ingin simpan ?',mtCustom,
                                  [mbYes,mbNo], 0)= mrNo
      then Exit ;

      simpandata;

      refreshdata;
   except
     ShowMessage('Gagal Simpan');
     // xRollback(frmMenu.conn);
     Exit;
   end;
    // xCommit(frmMenu.conn);
end;

procedure TfrmSubBarangPF.cxButton8Click(Sender: TObject);
begin
Release;
end;

procedure TfrmSubBarangPF.cxButton2Click(Sender: TObject);
begin
   try
      if (FLAGEDIT) and ( not cekedit(frmMenu.KDUSER,self.name)) then
        begin
           MessageDlg('Anda tidak berhak Edit di Modul ini',mtWarning, [mbOK],0);
           Exit;
        End;
         if (not FLAGEDIT) and ( not cekinsert(frmMenu.KDUSER,self.name)) then
        begin
           MessageDlg('Anda tidak berhak Insert di Modul ini',mtWarning, [mbOK],0);;
           Exit;
        End;

      if MessageDlg('Yakin ingin simpan ?',mtCustom,
                                  [mbYes,mbNo], 0)= mrNo
      then Exit ;

      simpandata;
      refreshdata;
   except
     ShowMessage('Gagal Simpan');
     // xRollback(frmMenu.conn);
     Exit;
   end;
    // xCommit(frmMenu.conn);
    Release;
end;

procedure TfrmSubBarangPF.edtKodeExit(Sender: TObject);
begin
loaddata(edtKode.Text);
end;

function TfrmSubBarangPF.GetCDSSubBarang: TClientDataset;
var s:String;
begin
  If not Assigned(FCDSSubBarang) then
  begin
    S := 'select nama_deptpf as Nama, kode_deptpf Kode'
        +' from tdeptpf ';


    FCDSSubBarang := TConextMain.cOpenCDS(S,nil);
  end;
  Result := FCDSSubBarang;
end;

procedure TfrmSubBarangPF.initview;
begin
  with TcxExtLookupHelper(cxLookupSubKode.Properties) do
    LoadFromCDS(CDSSubBarang, 'Kode','Nama',['Kode'],Self);
end;


end.
