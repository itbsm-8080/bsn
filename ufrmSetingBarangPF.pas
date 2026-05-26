unit ufrmSetingBarangPF;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, AdvPanel, ComCtrls, StdCtrls, AdvEdit,SqlExpr, Menus,
  cxLookAndFeelPainters, cxButtons,StrUtils, cxGraphics, cxLookAndFeels,
  dxSkinsCore, dxSkinsDefaultPainters, Grids, BaseGrid, AdvGrid, AdvCGrid,
  DBClient, cxStyles, cxCustomData, cxFilter, cxData, cxDataStorage,
  cxEdit, DB, cxDBData, cxSpinEdit, cxButtonEdit, cxTextEdit, cxGridLevel,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxClasses,
  cxControls, cxGridCustomView, cxGrid, AdvEdBtn, AdvCombo, cxCurrencyEdit,
  DateUtils, MyAccess;

type
  TfrmSetingBarangPF = class(TForm)
    AdvPanel1: TAdvPanel;
    AdvPanel3: TAdvPanel;
    cxButton8: TcxButton;
    AdvPanel2: TAdvPanel;
    lbljudul: TLabel;
    cxButton2: TcxButton;
    cxButton1: TcxButton;
    AdvPanel4: TAdvPanel;
    cxGrid: TcxGrid;
    cxGrdMain: TcxGridDBTableView;
    clNo: TcxGridDBColumn;
    clKode: TcxGridDBColumn;
    clNamaCustomer: TcxGridDBColumn;
    cxGrdDetail: TcxGridDBTableView;
    lvMaster: TcxGridLevel;
    clHet: TcxGridDBColumn;
    Label1: TLabel;
    Label4: TLabel;
    cbbBulan: TAdvComboBox;
    edtTahun: TComboBox;
    clnamagroup: TcxGridDBColumn;
    clkodegroup: TcxGridDBColumn;
    cxStyleRepository1: TcxStyleRepository;
    cxStyle1: TcxStyle;
    cxGrdMainColumn1: TcxGridDBColumn;
    clHNA: TcxGridDBColumn;
    procedure FormCreate(Sender: TObject);
    procedure refreshdata;
    procedure initgrid;
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure simpandata;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure cxButton1Click(Sender: TObject);
    procedure cxButton8Click(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    function GetCDS: TClientDataSet;
    procedure clNoGetDisplayText(Sender: TcxCustomGridTableItem; ARecord:
        TcxCustomGridRecord; var AText: string);
    procedure cxGrdMainKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure clKodePropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure clnamagroupPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
   procedure loaddata (abulan:String;atahun:string);      

  private
    FFLAGEDIT: Boolean;
    FID: string;


    { Private declarations }
  protected
    FCDS: TClientDataSet;
  public
    property CDS: TClientDataSet read GetCDS write FCDS;
    property FLAGEDIT: Boolean read FFLAGEDIT write FFLAGEDIT;
    property ID: string read FID write FID;
    { Public declarations }
  end;

var
  frmSetingBarangPF: TfrmSetingBarangPF;

implementation
uses MAIN,uModuleConnection,uFrmbantuan,Ulib;

{$R *.dfm}

procedure TfrmSetingBarangPF.FormCreate(Sender: TObject);
begin
     TcxDBGridHelper(cxGrdMain).LoadFromCDS(CDS, False, False);
end;

procedure TfrmSetingBarangPF.refreshdata;
begin
  FID:='';
  edtTahun.Text := FormatDateTime('yyyy',Date);
  initgrid;
end;
procedure TfrmSetingBarangPF.initgrid;
begin
  CDS.EmptyDataSet;
  CDS.Append;
  CDS.Post;
end;

procedure TfrmSetingBarangPF.FormKeyDown(Sender: TObject; var Key: Word;
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

procedure TfrmSetingBarangPF.FormKeyPress(Sender: TObject; var Key: Char);
begin
   if Key = #13 then
      SelectNext(ActiveControl,True,True);
end;


procedure TfrmSetingBarangPF.simpandata;
var
  s:string;
  i:integer;
  tt:TStrings;
  fid : integer;
begin


   s:= ' delete from tsetingbarangpf'
      + ' where  set_periode =' + inttostr(cbbBulan.ItemIndex+1)
      + ' and set_tahun = ' + edtTahun.Text;
      // xExecQuery(s,frmMenu.conn);
EnsureConnected(frmMenu.conn);
ExecSQLDirect(frmMenu.conn, s);
       // xCommit(frmMenu.conn);

   tt := TStringList.Create;
    CDS.First;
    i:=1;
  while not CDS.Eof do
  begin
   if not CDS.FieldByName('kode').IsNull then
   begin
    s:='insert into tsetingbarangpf (set_brg_kode,set_kode_grouppf,set_brg_het,set_periode,set_tahun,set_brg_hna)'
    + ' values ('
      + quot(CDS.FieldByName('kode').Asstring) + ','
      + quot(CDS.FieldByName('kodegroup').Asstring) + ','
      + FloatToStr(CDS.FieldByName('het').AsFloat) + ','
      + inttostr(cbbBulan.ItemIndex+1) + ','
      + edttahun.text+','
      + FloatToStr(CDS.FieldByName('hna').AsFloat)       
      + ');';
    tt.Append(s);
   end;
    CDS.Next;
    Inc(i);
  end;

     try
        for i:=0 to tt.Count -1 do
        begin
            // xExecQuery(tt[i],frmMenu.conn);
EnsureConnected(frmMenu.conn);
ExecSQLDirect(frmMenu.conn, tt[i]);
        end;
      finally
        tt.Free;
      end;
end;


procedure TfrmSetingBarangPF.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   Action := caFree;
   Release;
end;

procedure TfrmSetingBarangPF.cxButton1Click(Sender: TObject);
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

procedure TfrmSetingBarangPF.cxButton8Click(Sender: TObject);
begin
Release;
end;

procedure TfrmSetingBarangPF.cxButton2Click(Sender: TObject);
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
procedure TfrmSetingBarangPF.FormShow(Sender: TObject);
begin
refreshdata;
end;

function TfrmSetingBarangPF.GetCDS: TClientDataSet;
begin
  If not Assigned(FCDS) then
  begin
    FCDS := TClientDataSet.Create(Self);
    zAddField(FCDS, 'No', ftInteger, False);
    zAddField(FCDS, 'Kode', ftString, False,20);
    zAddField(FCDS, 'Nama', ftstring, False,100);
    zAddField(FCDS, 'HET', ftFloat, False);
    zAddField(FCDS, 'namaGroup', ftstring, False,100);
    zAddField(FCDS, 'kodegroup', ftInteger, False);
    zAddField(FCDS, 'HNA', ftFloat, False);
    FCDS.CreateDataSet;
  end;
  Result := FCDS;
end;

procedure TfrmSetingBarangPF.clNoGetDisplayText(Sender: TcxCustomGridTableItem;
    ARecord: TcxCustomGridRecord; var AText: string);
begin
  inherited;
  If Assigned(ARecord) then
  begin
    AText := Inttostr(ARecord.Index+1);
  end;
end;

procedure TfrmSetingBarangPF.cxGrdMainKeyDown(Sender: TObject; var Key: Word;
    Shift: TShiftState);
begin
if key = VK_DELETE then
begin
  If CDS.Eof then exit;
  CDS.Delete;
  If CDS.Eof then initgrid;
end;
end;

procedure TfrmSetingBarangPF.clKodePropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
  var
    i:integer;
begin
     sqlbantuan := ' SELECT brg_kode Kode,brg_nama Nama from tbarang order by brg_nama';
  sqlfilter := 'Kode,Nama';
  Application.CreateForm(Tfrmbantuan,frmbantuan);
  frmBantuan.SQLMaster := SQLbantuan;
  frmBantuan.ShowModal;
  if varglobal <> '' then
  begin
   for i := 0 to cxGrdMain.DataController.RecordCount-1 do
    begin

      If (VarToStr(cxGrdMain.DataController.Values[i, clkode.Index]) = VarToStr(varglobal)) and (cxGrdMain.DataController.FocusedRecordIndex <> i)
       then
      begin
          ShowMessage('Barang ada yang sama dengan baris '+ IntToStr(i+1));
          CDS.Cancel;
          exit;
      end;
    end;
   If CDS.State <> dsEdit then
         CDS.Edit;

      CDS.FieldByName('kode').AsString := varglobal;
      CDS.FieldByName('nama').AsString := varglobal1;

  end;

end;


procedure TfrmSetingBarangPF.clnamagroupPropertiesButtonClick(
  Sender: TObject; AButtonIndex: Integer);
  var
    i:integer;
begin
  sqlbantuan := ' SELECT  kode_grouppf Kode,nama_Grouppf Nama from tgrouppf order by nama_grouppf';
  sqlfilter := 'Kode,Nama';
  Application.CreateForm(Tfrmbantuan,frmbantuan);
  frmBantuan.SQLMaster := SQLbantuan;
  frmBantuan.ShowModal;
  if varglobal <> '' then
  begin

   If CDS.State <> dsEdit then
         CDS.Edit;

      CDS.FieldByName('kodeGroup').AsString := varglobal;
      CDS.FieldByName('namaGroup').AsString := varglobal1;

  end;

end;

procedure TfrmSetingBarangPF.loaddata (abulan:String;atahun:string);
var
  s:string;
  tsql:TmyQuery;
begin
 s:='select set_brg_kode,brg_nama,set_brg_het,set_kode_grouppf,nama_grouppf,set_brg_hna'
  + ' from tsetingbarangpf inner join tbarang on brg_kode=set_brg_kode '
  + ' inner join tgrouppf on set_kode_grouppf=kode_grouppf '
  + ' where set_periode = '+abulan
  + ' and set_tahun = ' + atahun ;
  tsql := xOpenQuery(s,frmMenu.conn);
 with tsql do
 begin
   try
   tsql.First;
   cds.EmptyDataSet;
   while not Eof do
   begin
     cds.Append;
     cds.FieldByName('kode').AsString := Fields[0].AsString;
     cds.FieldByName('nama').AsString := Fields[1].AsString;
     cds.FieldByName('het').Asfloat := Fields[2].AsFloat;
     cds.FieldByName('kodegroup').AsString := Fields[3].AsString;
     cds.FieldByName('namagroup').AsString := Fields[4].AsString;
     cds.FieldByName('hna').Asfloat := Fields[5].AsFloat;
     cds.post;
     Next;
   end;
   finally
     free;
   end;
 end;


end;

end.
