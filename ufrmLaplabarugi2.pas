unit ufrmLapLabaRugi2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids,  AdvGrid, ComCtrls, StdCtrls, AdvEdit, ExtCtrls,
  AdvPanel, AdvCGrid, BaseGrid,SqlExpr, DBAdvGrd, DB, DBClient, Provider,
  FMTBcd, RAWPrinter, StrUtils, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, dxSkinsCore,
  dxSkinsDefaultPainters, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxLookupEdit, cxDBLookupEdit, cxDBExtLookupComboBox, cxStyles,
  dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage,
  cxDBData, cxSpinEdit, cxCalendar, Menus, cxButtons, cxGridLevel,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxClasses,
  cxGridCustomView, cxGrid, cxButtonEdit, cxCurrencyEdit,ExcelXP,ComObj,
  AdvCombo,DateUtils, cxPC, cxGridBandedTableView, cxGridDBBandedTableView,
  MyAccess, MemDS, DBAccess;

type
  TfrmLapLabaRugi2 = class(TForm)
    AdvPanel1: TAdvPanel;
    lbljudul: TLabel;
    AdvPanel2: TAdvPanel;
    AdvPanel3: TAdvPanel;
    Label3: TLabel;
    RAWPrinter1: TRAWPrinter;
    AdvPanel4: TAdvPanel;
    cxButton8: TcxButton;
    PopupMenu1: TPopupMenu;
    HapusRecord1: TMenuItem;
    OpenDialog1: TOpenDialog;
    edtTahun: TComboBox;
    Button2: TButton;
    cxStyleRepository2: TcxStyleRepository;
    cxStyle2: TcxStyle;
    MainMenu1: TMainMenu;
    cxButton7: TcxButton;
    SaveDialog1: TSaveDialog;
    cxGrid1: TcxGrid;
    cxGrid1DBBandedTableView1: TcxGridDBBandedTableView;
    clKeterangan: TcxGridDBBandedColumn;
    clBulanini: TcxGridDBBandedColumn;
    clbulaninithlalu: TcxGridDBBandedColumn;
    clMei: TcxGridDBBandedColumn;
    clJuli: TcxGridDBBandedColumn;
    clApr: TcxGridDBBandedColumn;
    clAgustus: TcxGridDBBandedColumn;
    clratiobulanini: TcxGridDBBandedColumn;
    cxGrid1Level1: TcxGridLevel;
    clSept: TcxGridDBBandedColumn;
    clNov: TcxGridDBBandedColumn;
    clDes: TcxGridDBBandedColumn;
    cxStyleRepository1: TcxStyleRepository;
    cxStyle1: TcxStyle;
    clTotal: TcxGridDBBandedColumn;
    dtstprvdr1: TDataSetProvider;
    ds2: TDataSource;
    ds3: TClientDataSet;
    sqlqry2: TSQLQuery;
    cxStyle3: TcxStyle;
    Label1: TLabel;
    cbbBulan: TAdvComboBox;
    cxStyle4: TcxStyle;
    sqlqry1: TMyQuery;
    procedure Button2Click(Sender: TObject);
    procedure loaddata ;
    function gettotal(abulan:integer;akol:string) : double;
    function gettotal2(abulan:integer;akol:string) : double;
    function gettotal3(abulan:integer;akol:string) : double;
function gettotal4(abulan:integer;akol:string) : double;
function gettotal5(abulan:integer;akol:string) : double;    
    procedure cxGrid1DBBandedTableView1StylesGetContentStyle(
      Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
      AItem: TcxCustomGridTableItem; out AStyle: TcxStyle);
    procedure cxButton7Click(Sender: TObject);
    procedure cxButton8Click(Sender: TObject);
    private
     buttonSelected  : integer;
     FID : STRING;
     FFLAGEDIT: Boolean;
     xtotal : Double;
      { Private declarations }
     protected
  public
    property ID: string read FID write FID;
    property FLAGEDIT: Boolean read FFLAGEDIT write FFLAGEDIT;
    { Public declarations }
  end;

var
  frmLapLabaRugi2: TfrmLapLabaRugi2;

implementation

uses MAIN,uModuleConnection,uFrmbantuan,Ulib,uReport,cxgridExportlink;

{$R *.dfm}


procedure TfrmLapLabaRugi2.Button2Click(Sender: TObject);
begin
loaddata;
end;

procedure TfrmLapLabaRugi2.loaddata ;
var
  ssql :string;
  tsql : TmyQuery;
  i:integer;
  atotal :Double;
begin
  ssql := 'select 1 no,rek_nama Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok'
          + ' on kol_id=rek_kol_id'
          + ' where kol_id=10'
          + ' union'
          + ' select distinct 2 no,"TOTAL PENJUALAN" Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok'
          + ' on kol_id=rek_kol_id'
          + ' where kol_id=10 '
          + ' union'
          + ' select 3 no,rek_nama Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok'
          + ' on kol_id=rek_kol_id'
          + ' where kol_id=9'
          + ' union'
          + ' select distinct 4 no,"LABA KOTOR" Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok'
          + ' on kol_id=rek_kol_id'
          + ' where kol_id=10 '
          + ' union'
          + ' select distinct 5 no,"RATIO 1" Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok'
          + ' on kol_id=rek_kol_id'
          + ' where kol_id=10 '

          + ' union'
          + ' select * from (select 6 no,rek_nama Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok'
          + ' on kol_id=rek_kol_id'
          + ' where kol_id in (12,13) order by rek_urutan ) a '
          + ' union'
          + ' select distinct 7 no ,"TOTAL BIAYA" Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok'
          + ' on kol_id=rek_kol_id'
          + ' where kol_id=10 '
          + ' union'
          + ' select distinct 8 no,"RATIO 2" Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok'
          + ' on kol_id=rek_kol_id'
          + ' where kol_id=10 '
          + ' union'
          + ' select 9 no,rek_nama Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok'
          + ' on kol_id=rek_kol_id'
          + ' where kol_id in (11,14)'
          + ' union'
          + ' select distinct 10 no,"TOTAL PENDAPATAN" Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok '
          + ' on kol_id=rek_kol_id'
          + ' where kol_id=9 '
          + ' union'
          + ' select distinct 11 no,"RATIO 3" Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok'
          + ' on kol_id=rek_kol_id'
          + ' where kol_id=10 '

          + ' union'

          + ' select 12 no,"LABA BERSIH" Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok'
          + ' on kol_id=rek_kol_id'
          + ' where kol_id=9'
          + ' union'

          + ' select 13 no,"RATIO 4" Keterangan,'
          + ' cast(0 as decimal) Bulanini,cast(0 as decimal) bulaninilalu,cast(0 as decimal) Ratio,'
          + ' cast(0 as decimal) sdbulanini,cast(0 as decimal) sdbulaninilalu,cast(0 as decimal) Ratio2, '
          + ' cast(0 as decimal) sdbulanini2,cast(0 as decimal) tahunlalufull,cast(0 as decimal) Ratio3'
          + ' from trekening inner join tkelompok'
          + ' on kol_id=rek_kol_id'
          + ' where kol_id=9'

          ;
        ds3.Close;
        sqlqry1.Connection := frmMenu.conn;
        sqlqry1.SQL.Text := ssql;
        ds3.open;
// bulan ini
        ssql := 'select month(jur_tanggal) bulan,'
            + ' rek_nama keterangan,sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+edtTahun.Text
            + ' and month(jur_tanggal)='+ IntToStr(cbbBulan.ItemIndex+1)
            + ' group by month(jur_tanggal) , '
            + ' rek_kode,kol_id '
            + ' order by month(jur_tanggal) ,kol_id,rek_nama';

        tsql :=xOpenQuery(ssql,frmMenu.conn);
        with tsql do
        begin
          try
            while not Eof do
            begin
              if ds3.Locate('keterangan',FieldByName('keterangan').AsString,[loCaseInsensitive]) then
              begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulanini').AsFloat := FieldByName('nilai').AsFloat/1000;
              end;
              Next;
            end;
            if ds3.Locate('keterangan','TOTAL PENJUALAN',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulanini').AsFloat := gettotal(cbbBulan.ItemIndex+1,'10')/1000 ;
            end;
            if ds3.Locate('keterangan','LABA KOTOR',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulanini').AsFloat := gettotal(cbbBulan.ItemIndex+1 ,'10,9')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 1',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulanini').AsFloat := (gettotal(cbbBulan.ItemIndex+1 ,'10,9')/1000)/(gettotal(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;

//
            if ds3.Locate('keterangan','TOTAL BIAYA',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulanini').AsFloat := gettotal(cbbBulan.ItemIndex+1 ,'12,13')/1000 ;
            end;

            if ds3.Locate('keterangan','RATIO 2',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulanini').AsFloat := (gettotal(cbbBulan.ItemIndex+1 ,'12,13')/1000)/(gettotal(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;

            if ds3.Locate('keterangan','TOTAL PENDAPATAN',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulanini').AsFloat := gettotal(cbbBulan.ItemIndex+1 ,'11,14')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 3',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulanini').AsFloat := (gettotal(cbbBulan.ItemIndex+1 ,'11,14')/1000)/(gettotal(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;

            if ds3.Locate('keterangan','LABA BERSIH',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulanini').AsFloat := gettotal(cbbBulan.ItemIndex+1 ,'')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 4',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulanini').AsFloat := (gettotal(cbbBulan.ItemIndex+1 ,'')/1000 )/(gettotal(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;


           finally
             free;
           end;
         end;
// bulan ini tahun lalu

        ssql := 'select month(jur_tanggal) bulan,'
            + ' rek_nama keterangan,sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+inttostr(StrToInt(edtTahun.Text)-1)
            + ' and month(jur_tanggal)='+ IntToStr(cbbBulan.ItemIndex+1)
            + ' group by month(jur_tanggal) , '
            + ' rek_kode,kol_id '
            + ' order by month(jur_tanggal) ,kol_id,rek_nama';

        tsql :=xOpenQuery(ssql,frmMenu.conn);
        with tsql do
        begin
          try
            while not Eof do
            begin
              if ds3.Locate('keterangan',FieldByName('keterangan').AsString,[loCaseInsensitive]) then
              begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulaninilalu').AsFloat := FieldByName('nilai').AsFloat/1000;
              end;
              Next;
            end;
            if ds3.Locate('keterangan','TOTAL PENJUALAN',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulaninilalu').AsFloat := gettotal2(cbbBulan.ItemIndex+1,'10')/1000 ;
            end;
            if ds3.Locate('keterangan','LABA KOTOR',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulaninilalu').AsFloat := gettotal2(cbbBulan.ItemIndex+1 ,'10,9')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 1',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulaninilalu').AsFloat := (gettotal2(cbbBulan.ItemIndex+1 ,'10,9')/1000)/(gettotal2(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;


//
            if ds3.Locate('keterangan','TOTAL BIAYA',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulaninilalu').AsFloat := gettotal2(cbbBulan.ItemIndex+1 ,'12,13')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 2',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulaninilalu').AsFloat := (gettotal2(cbbBulan.ItemIndex+1 ,'12,13')/1000)/(gettotal2(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;

            if ds3.Locate('keterangan','TOTAL PENDAPATAN',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulaninilalu').AsFloat := gettotal2(cbbBulan.ItemIndex+1 ,'11,14')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 3',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulaninilalu').AsFloat := (gettotal2(cbbBulan.ItemIndex+1 ,'11,14')/1000)/(gettotal2(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;

            if ds3.Locate('keterangan','LABA BERSIH',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulaninilalu').AsFloat := gettotal2(cbbBulan.ItemIndex+1 ,'')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 4',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('bulaninilalu').AsFloat := (gettotal2(cbbBulan.ItemIndex+1 ,'')/1000)/(gettotal2(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;


           finally
             free;
           end;
         end;

// sd bulan ini

        ssql := 'select '
            + ' rek_nama keterangan,sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+edtTahun.Text
            + ' and month(jur_tanggal) between 1 and '+ IntToStr(cbbBulan.ItemIndex+1)
            + ' group by year(jur_tanggal) , '
            + ' rek_kode,kol_id '
            + ' order by kol_id,rek_nama';

        tsql :=xOpenQuery(ssql,frmMenu.conn);
        with tsql do
        begin
          try
            while not Eof do
            begin
              if ds3.Locate('keterangan',FieldByName('keterangan').AsString,[loCaseInsensitive]) then
              begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulanini').AsFloat := FieldByName('nilai').AsFloat/1000;
                ds3.FieldByName('sdbulanini2').AsFloat := FieldByName('nilai').AsFloat/1000;
              end;
              Next;
            end;
            if ds3.Locate('keterangan','TOTAL PENJUALAN',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulanini').AsFloat := gettotal3(cbbBulan.ItemIndex+1,'10')/1000 ;
                ds3.FieldByName('sdbulanini2').AsFloat := gettotal3(cbbBulan.ItemIndex+1,'10')/1000 ;

            end;
            if ds3.Locate('keterangan','LABA KOTOR',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulanini').AsFloat := gettotal3(cbbBulan.ItemIndex+1 ,'10,9')/1000 ;
                ds3.FieldByName('sdbulanini2').AsFloat := gettotal3(cbbBulan.ItemIndex+1 ,'10,9')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 1',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulanini').AsFloat := (gettotal3(cbbBulan.ItemIndex+1 ,'10,9')/1000)/(gettotal3(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
                ds3.FieldByName('sdbulanini2').AsFloat := (gettotal3(cbbBulan.ItemIndex+1 ,'10,9')/1000)/(gettotal3(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;


//
            if ds3.Locate('keterangan','TOTAL BIAYA',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulanini').AsFloat := gettotal3(cbbBulan.ItemIndex+1 ,'12,13')/1000 ;
                ds3.FieldByName('sdbulanini2').AsFloat := gettotal3(cbbBulan.ItemIndex+1 ,'12,13')/1000 ;
            end;

            if ds3.Locate('keterangan','RATIO 2',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulanini').AsFloat := (gettotal3(cbbBulan.ItemIndex+1 ,'12,13')/1000)/(gettotal3(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
                ds3.FieldByName('sdbulanini2').AsFloat := (gettotal3(cbbBulan.ItemIndex+1 ,'12,13')/1000)/(gettotal3(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;


            if ds3.Locate('keterangan','TOTAL PENDAPATAN',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulanini').AsFloat := gettotal3(cbbBulan.ItemIndex+1 ,'11,14')/1000 ;
                ds3.FieldByName('sdbulanini2').AsFloat := gettotal3(cbbBulan.ItemIndex+1 ,'11,14')/1000 ;

            end;
            if ds3.Locate('keterangan','RATIO 3',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulanini').AsFloat := (gettotal3(cbbBulan.ItemIndex+1 ,'11,14')/1000)/(gettotal3(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
                ds3.FieldByName('sdbulanini2').AsFloat := (gettotal3(cbbBulan.ItemIndex+1 ,'11,14')/1000)/(gettotal3(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;


            if ds3.Locate('keterangan','LABA BERSIH',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulanini').AsFloat := gettotal3(cbbBulan.ItemIndex+1 ,'')/1000 ;
                ds3.FieldByName('sdbulanini2').AsFloat := gettotal3(cbbBulan.ItemIndex+1 ,'')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 4',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulanini').AsFloat := (gettotal3(cbbBulan.ItemIndex+1 ,'')/1000)/(gettotal3(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
                ds3.FieldByName('sdbulanini2').AsFloat := (gettotal3(cbbBulan.ItemIndex+1 ,'')/1000)/(gettotal3(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;


           finally
             free;
           end;
         end;


// sd bulan inilalu

        ssql := 'select '
            + ' rek_nama keterangan,sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+inttostr(StrToInt(edtTahun.Text)-1)
            + ' and month(jur_tanggal) between 1 and '+ IntToStr(cbbBulan.ItemIndex+1)
            + ' group by year(jur_tanggal) , '
            + ' rek_kode,kol_id '
            + ' order by kol_id,rek_nama';

        tsql :=xOpenQuery(ssql,frmMenu.conn);
        with tsql do
        begin
          try
            while not Eof do
            begin
              if ds3.Locate('keterangan',FieldByName('keterangan').AsString,[loCaseInsensitive]) then
              begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulaninilalu').AsFloat := FieldByName('nilai').AsFloat/1000;
              end;
              Next;
            end;
            if ds3.Locate('keterangan','TOTAL PENJUALAN',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulaninilalu').AsFloat := gettotal4(cbbBulan.ItemIndex+1,'10')/1000 ;
            end;
            if ds3.Locate('keterangan','LABA KOTOR',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulaninilalu').AsFloat := gettotal4(cbbBulan.ItemIndex+1 ,'10,9')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 1',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulaninilalu').AsFloat := (gettotal4(cbbBulan.ItemIndex+1 ,'10,9')/1000)/(gettotal4(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;

//
            if ds3.Locate('keterangan','TOTAL BIAYA',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulaninilalu').AsFloat := gettotal4(cbbBulan.ItemIndex+1 ,'12,13')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 2',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulaninilalu').AsFloat := (gettotal4(cbbBulan.ItemIndex+1 ,'12,13')/1000)/(gettotal4(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;

            if ds3.Locate('keterangan','TOTAL PENDAPATAN',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulaninilalu').AsFloat := gettotal4(cbbBulan.ItemIndex+1 ,'11,14')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 3',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulaninilalu').AsFloat := (gettotal4(cbbBulan.ItemIndex+1 ,'11,14')/1000)/(gettotal4(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;

            if ds3.Locate('keterangan','LABA BERSIH',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulaninilalu').AsFloat := gettotal4(cbbBulan.ItemIndex+1 ,'')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 4',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('sdbulaninilalu').AsFloat := (gettotal4(cbbBulan.ItemIndex+1 ,'')/1000)/(gettotal4(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;



           finally
             free;
           end;
         end;

// tahun lalu

        ssql := 'select '
            + ' rek_nama keterangan,sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+inttostr(StrToInt(edtTahun.Text)-1)
            + ' group by year(jur_tanggal) , '
            + ' rek_kode,kol_id '
            + ' order by kol_id,rek_nama';

        tsql :=xOpenQuery(ssql,frmMenu.conn);
        with tsql do
        begin
          try
            while not Eof do
            begin
              if ds3.Locate('keterangan',FieldByName('keterangan').AsString,[loCaseInsensitive]) then
              begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('tahunlalufull').AsFloat := FieldByName('nilai').AsFloat/1000;
              end;
              Next;
            end;
            if ds3.Locate('keterangan','TOTAL PENJUALAN',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('tahunlalufull').AsFloat := gettotal5(cbbBulan.ItemIndex+1,'10')/1000 ;
            end;
            if ds3.Locate('keterangan','LABA KOTOR',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('tahunlalufull').AsFloat := gettotal5(cbbBulan.ItemIndex+1 ,'10,9')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 1',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('tahunlalufull').AsFloat := (gettotal5(cbbBulan.ItemIndex+1 ,'10,9')/1000)/(gettotal5(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;

//
            if ds3.Locate('keterangan','TOTAL BIAYA',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('tahunlalufull').AsFloat := gettotal5(cbbBulan.ItemIndex+1 ,'12,13')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 2',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('tahunlalufull').AsFloat := (gettotal5(cbbBulan.ItemIndex+1 ,'12,13')/1000)/(gettotal5(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;

            if ds3.Locate('keterangan','TOTAL PENDAPATAN',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('tahunlalufull').AsFloat := gettotal5(cbbBulan.ItemIndex+1 ,'11,14')/1000 ;
            end;
            if ds3.Locate('keterangan','RATIO 3',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('tahunlalufull').AsFloat := (gettotal5(cbbBulan.ItemIndex+1 ,'11,14')/1000)/(gettotal5(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;

            if ds3.Locate('keterangan','LABA BERSIH',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('tahunlalufull').AsFloat := gettotal5(cbbBulan.ItemIndex+1 ,'')/1000 ;
            end;

            if ds3.Locate('keterangan','RATIO 4',[loCaseInsensitive]) then
            begin
                If ds3.State <> dsEdit then ds3.Edit;
                ds3.FieldByName('tahunlalufull').AsFloat := (gettotal5(cbbBulan.ItemIndex+1 ,'')/1000)/(gettotal5(cbbBulan.ItemIndex+1,'10')/1000)*100 ;
            end;

           finally
             free;
           end;
         end;
         ds3.First;
         while not ds3.Eof do
         begin
            If ds3.State <> dsEdit then ds3.Edit;
               if (ds3.FieldByName('bulaninilalu').AsFloat <> 0) and (copy(ds3.FieldByName('keterangan').AsString,1,5) <> 'RATIO') then
               ds3.FieldByName('ratio').AsFloat := ds3.FieldByName('bulanini').AsFloat /ds3.FieldByName('bulaninilalu').AsFloat*100 ;
               if (ds3.FieldByName('sdbulaninilalu').AsFloat <> 0) and (copy(ds3.FieldByName('keterangan').AsString,1,5) <> 'RATIO') then
               ds3.FieldByName('ratio2').AsFloat := ds3.FieldByName('sdbulanini').AsFloat /ds3.FieldByName('sdbulaninilalu').AsFloat*100 ;
               if (ds3.FieldByName('tahunlalufull').AsFloat <> 0) and (copy(ds3.FieldByName('keterangan').AsString,1,5) <> 'RATIO') then
               ds3.FieldByName('ratio3').AsFloat := ds3.FieldByName('sdbulanini').AsFloat /ds3.FieldByName('tahunlalufull').AsFloat*100 ;

           ds3.Next;
         end;

end;

function TfrmLapLabaRugi2.gettotal(abulan:integer;akol:string) : double;
var
  s:string;
  tsql :TmyQuery;
begin
  Result := 0;
  IF AKOL <> ''  Then
        s := 'select '
            + ' sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+edtTahun.Text
            + ' and month(jur_tanggal)='+ IntToStr(abulan)
            + ' and kol_id  in ('+akol+')'
  else
          s := 'select '
            + ' sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+edtTahun.Text
            + ' and month(jur_tanggal)='+ IntToStr(abulan);


            tsql := xOpenQuery(s,frmMenu.conn);
      with tsql do
      begin
        if not eof then
           Result := Fields[0].AsFloat;
      end;

end;
function TfrmLapLabaRugi2.gettotal2(abulan:integer;akol:string) : double;
var
  s:string;
  tsql :TmyQuery;
begin
  Result := 0;
  IF AKOL <> ''  Then
        s := 'select '
            + ' sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+inttostr(StrToInt(edtTahun.Text)-1)
            + ' and month(jur_tanggal)='+ IntToStr(abulan)
            + ' and kol_id  in ('+akol+')'
  else
          s := 'select '
            + ' sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+inttostr(StrToInt(edtTahun.Text)-1)
            + ' and month(jur_tanggal)='+ IntToStr(abulan);


            tsql := xOpenQuery(s,frmMenu.conn);
      with tsql do
      begin
        if not eof then
           Result := Fields[0].AsFloat;
      end;

end;

function TfrmLapLabaRugi2.gettotal3(abulan:integer;akol:string) : double;
var
  s:string;
  tsql :TmyQuery;
begin
  Result := 0;
  IF AKOL <> ''  Then
        s := 'select '
            + ' sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+edtTahun.Text
            + ' and month(jur_tanggal) between 1 and '+ IntToStr(abulan)
            + ' and kol_id  in ('+akol+')'
  else
          s := 'select '
            + ' sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+edtTahun.Text
            + ' and month(jur_tanggal) between 1 and '+ IntToStr(abulan);


            tsql := xOpenQuery(s,frmMenu.conn);
      with tsql do
      begin
        if not eof then
           Result := Fields[0].AsFloat;
      end;

end;

function TfrmLapLabaRugi2.gettotal4(abulan:integer;akol:string) : double;
var
  s:string;
  tsql :TmyQuery;
begin
  Result := 0;
  IF AKOL <> ''  Then
        s := 'select '
            + ' sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+inttostr(StrToInt(edtTahun.Text)-1)
            + ' and month(jur_tanggal) between 1 and '+ IntToStr(abulan)
            + ' and kol_id  in ('+akol+')'
  else
          s := 'select '
            + ' sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+inttostr(StrToInt(edtTahun.Text)-1)
            + ' and month(jur_tanggal) between 1 and '+ IntToStr(abulan);


            tsql := xOpenQuery(s,frmMenu.conn);
      with tsql do
      begin
        if not eof then
           Result := Fields[0].AsFloat;
      end;

end;

function TfrmLapLabaRugi2.gettotal5(abulan:integer;akol:string) : double;
var
  s:string;
  tsql :TmyQuery;
begin
  Result := 0;
  IF AKOL <> ''  Then
        s := 'select '
            + ' sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+inttostr(StrToInt(edtTahun.Text)-1)
            + ' and kol_id  in ('+akol+')'
  else
          s := 'select '
            + ' sum(ifnull(jurd_kredit,0)-ifnull(jurd_debet,0)) nilai from  tjurnal inner join tjurnalitem'
            + ' on jur_no =jurd_jur_no'
            + ' inner join trekening on rek_kode=jurd_rek_kode'
            + ' inner join tkelompok on kol_id=rek_kol_id'
            + ' where isneraca=0 and year(jur_tanggal)='+inttostr(StrToInt(edtTahun.Text)-1);


            tsql := xOpenQuery(s,frmMenu.conn);
      with tsql do
      begin
        if not eof then
           Result := Fields[0].AsFloat;
      end;

end;

procedure TfrmLapLabaRugi2.cxGrid1DBBandedTableView1StylesGetContentStyle(
  Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
  AItem: TcxCustomGridTableItem; out AStyle: TcxStyle);
begin
if (ARecord <> nil) and (AItem <> nil) and (VarToStr(ARecord.Values[0]) = 'TOTAL PENJUALAN') then
    AStyle := cxStyle2;
if (ARecord <> nil) and (AItem <> nil) and (VarToStr(ARecord.Values[0]) = 'LABA KOTOR') then
    AStyle := cxStyle2;
if (ARecord <> nil) and (AItem <> nil) and (COPY(VarToStr(ARecord.Values[0]),1,5) = 'RATIO') then
    AStyle := cxStyle4;
if (ARecord <> nil) and (AItem <> nil) and (VarToStr(ARecord.Values[0]) = 'TOTAL BIAYA') then
    AStyle := cxStyle2;
if (ARecord <> nil) and (AItem <> nil) and (VarToStr(ARecord.Values[0]) = 'TOTAL PENDAPATAN') then
    AStyle := cxStyle2;
if (ARecord <> nil) and (AItem <> nil) and (VarToStr(ARecord.Values[0]) = 'LABA BERSIH') then
    AStyle := cxStyle3;

end;

procedure TfrmLapLabaRugi2.cxButton7Click(Sender: TObject);
begin
  if SaveDialog1.Execute then
     begin
       ExportGridToExcel(SaveDialog1.FileName, cxGrid1);
     end;

end;

procedure TfrmLapLabaRugi2.cxButton8Click(Sender: TObject);
begin
Release;
end;

end.
