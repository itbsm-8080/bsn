unit ufrmBrowseSetingBarangPF;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ufrmCxBrowse, Menus, cxLookAndFeelPainters, cxStyles,
  dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee,
  dxSkinDarkSide,
  dxSkinGlassOceans, dxSkiniMaginary, dxSkinLilian,
  dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinPumpkin,
  dxSkinSilver, dxSkinSpringTime, dxSkinStardust,
  dxSkinSummer2008, dxSkinValentine, dxSkinXmas2008Blue,
  dxSkinscxPCPainter, cxCustomData, cxGraphics, cxFilter, cxData,
  cxDataStorage, cxEdit, DB, cxDBData, FMTBcd, Provider, SqlExpr, ImgList,
  ComCtrls, StdCtrls, cxGridLevel, cxClasses, cxControls, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  cxButtons, ExtCtrls, AdvPanel, DBClient, cxLookAndFeels;

type
  TfrmBrowseSetingBarangPF = class(TfrmCxBrowse)
    cxButton5: TcxButton;
  procedure btnRefreshClick(Sender: TObject);
  procedure FormShow(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
  procedure cxButton6Click(Sender: TObject);
    procedure cxButton4Click(Sender: TObject);
    procedure cxButton5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmBrowseSetingBarangPF: TfrmBrowseSetingBarangPF;

implementation
   uses ufrmSetingBarangPF,Ulib, MAIN, uModuleConnection;
{$R *.dfm}

procedure TfrmBrowseSetingBarangPF.btnRefreshClick(Sender: TObject);
begin
  Self.SQLMaster := 'select '
  + ' set_periode Periode,set_tahun Tahun,brg_kode Kode,brg_nama Nama,nama_grouppf Grup,set_brg_het Het,nama_deptpf Dept,set_brg_hna HNA'
  + '  from tsetingbarangpf'
  + ' inner join tgrouppf y  on kode_grouppf=set_kode_grouppf '
  + ' inner join tdeptpf x on x.kode_deptpf=y.kode_deptpf'
  + ' inner join tbarang on brg_kode=set_brg_kode';
   inherited;
    cxGrdMaster.ApplyBestFit();
    cxGrdMaster.Columns[0].Width :=100;
    cxGrdMaster.Columns[1].Width :=100;
end;

procedure TfrmBrowseSetingBarangPF.FormShow(Sender: TObject);
begin
    ShowWindowAsync(Handle, SW_MAXIMIZE);
  inherited;
  btnRefreshClick(Self);
end;

procedure TfrmBrowseSetingBarangPF.cxButton2Click(Sender: TObject);
var
  frmSetingBarangPF: TfrmSetingBarangPF;
begin
  inherited;
    if ActiveMDIChild.Caption <> 'Seting Barang PF' then
   begin
      frmSetingBarangPF  := frmmenu.ShowForm(TfrmSetingBarangPF) as TfrmSetingBarangPF;
   end;
   frmSetingBarangPF.Show;
end;

procedure TfrmBrowseSetingBarangPF.cxButton1Click(Sender: TObject);
var
  frmSetingBarangPF: TfrmSetingBarangPF;
begin
  inherited;
  If CDSMaster.FieldByname('periode').IsNull then exit;
  if ActiveMDIChild.Caption <> 'Seting Barang PF' then
   begin
//      ShowForm(TfrmBrowseBarang).Show;
      frmSetingBarangPF  := frmmenu.ShowForm(TfrmSetingBarangPF) as TfrmSetingBarangPF;
//      frmSetingBarangPF.ID := CDSMaster.FieldByname('KODE').AsString;
      frmSetingBarangPF.FLAGEDIT := True;
      frmSetingBarangPF.cbbBulan.ItemIndex := StrToInt(CDSMaster.FieldByname('Periode').AsString)-1;
      frmSetingBarangPF.edtTahun.Text := CDSMaster.FieldByname('Tahun').AsString;
      frmSetingBarangPF.loaddata(CDSMaster.FieldByname('periode').AsString,CDSMaster.FieldByname('tahun').AsString);

   end;
   frmSetingBarangPF.Show;
end;

procedure TfrmBrowseSetingBarangPF.cxButton6Click(Sender: TObject);
begin
  inherited;
  refreshdata;
end;

procedure TfrmBrowseSetingBarangPF.cxButton4Click(Sender: TObject);
var
  s:string;
begin
  inherited;
     try
       if not cekdelete(frmMenu.KDUSER,'frmGudang') then
      begin
         MessageDlg('Anda tidak berhak Menghapus di Modul ini',mtWarning, [mbOK],0);
         Exit;
      End;
      if MessageDlg('Yakin ingin hapus ?',mtCustom,
                                  [mbYes,mbNo], 0)= mrNo
      then Exit ;
       s:='delete from tsetingbarangpf '
        + ' where set_periode = ' + quot(CDSMaster.FieldByname('periode').AsString)
        + ' and set_tahun = ' + quot(CDSMaster.FieldByname('tahun').AsString)+';' ;
      // xExecQuery(s,frmMenu.conn);
EnsureConnected(frmMenu.conn);
ExecSQLDirect(frmMenu.conn, s);


//      CDSMaster.Delete;
   except
     MessageDlg('Gagal Hapus',mtError, [mbOK],0);
     // xRollback(frmMenu.conn);
     Exit;
   end;
    // xCommit(frmMenu.conn);
  btnRefreshClick(self);
end;

procedure TfrmBrowseSetingBarangPF.cxButton5Click(Sender: TObject);
var
  s:string;
  tt:TStrings;
begin
  inherited;
  if savedlg.Execute then
  begin
      CDSMaster.First;
      CDSMaster.Filter := cxGrdMaster.DataController.Filter.FilterText;
      CDSMaster.Filtered := True;
      tt := TStringList.Create;

      with CDSMaster do
      begin
        s:=' delete from tbarangpf where bpf_periode='
          + IntToStr(fieldbyname('periode').asinteger)+' and bpf_tahun='
          + IntToStr(fieldbyname('tahun').asinteger)+';';
          tt.Append(s);
        First;
        while not Eof do
        begin
           s:='insert into tbarangpf (bpf_periode,bpf_tahun,bpf_brg_kode,bpf_nama,bpf_grup,bpf_het,bpf_dept,bpf_hna)'
           + ' values ('
           + IntToStr(fieldbyname('periode').asinteger)+','
           + IntToStr(fieldbyname('tahun').asinteger)+','
           + IntToStr(fieldbyname('kode').asinteger)+','
           + quot(fieldbyname('nama').AsString)+','
           + Quot(fieldbyname('grup').AsString)+','
           + FloatToStr(fieldbyname('het').Asfloat)+','
           + Quot(fieldbyname('dept').AsString)+','
           + FloatToStr(fieldbyname('hna').Asfloat)+');';
           tt.Append(s);

           Next;
        end;
      end;
      tt.SaveToFile(savedlg.FileName);
  end;

end;

end.
