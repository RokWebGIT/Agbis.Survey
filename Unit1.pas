unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.FB,
  FireDAC.Phys.FBDef, FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, Vcl.StdCtrls, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, Vcl.ExtCtrls, Vcl.WinXPickers, Vcl.CheckLst,
  FireDAC.VCLUI.Async, FireDAC.Comp.UI, FireDAC.Phys.IBWrapper, StrUtils, DateUtils, ComObj,
  FireDAC.Phys.IBDef, FireDAC.Phys.IBBase, FireDAC.Phys.IB, FireDAC.FMXUI.Async,
  Vcl.Menus;

type
  TForm1 = class(TForm)
    OpenDialog1: TOpenDialog;
    FDGUIxAsyncExecuteDialog1: TFDGUIxAsyncExecuteDialog;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  Connection: TFDConnection;
  Query: TFDQuery;

implementation

{$R *.dfm}

procedure TForm1.FormShow(Sender: TObject);
var
  Excel:Variant;
  I: integer;
  PhoneColumn: Integer;
  PhoneRowStart: Integer;
  PhoneRowEnd: Integer;
  Phone: String;
  const xlCellTypeLast=$000000B;
begin
  Excel := CreateoleObject('Excel.Application');
  try
    OpenDialog1.Filter:='Файлы MS Excel|*.xls;*.xlsx|';
    if  not OpenDialog1.Execute then exit;

    Excel.Workbooks.Open(OpenDialog1.FileName, False);
    Excel.Visible := true;
    Excel.Cells[1,1].Select;
    Excel.Cells.Find(What := 'Open-Ended Response', After := Excel.ActiveCell,
                     LookIn := -4163, LookAt := 1,
                     SearchOrder := 1, SearchDirection := 1,
                     MatchCase := false, SearchFormat := false).Activate;
    PhoneColumn := Excel.ActiveCell.Column;
    PhoneRowStart := Excel.ActiveCell.Row+1;
    PhoneRowEnd := Excel.ActiveCell.SpecialCells(xlCellTypeLast).Row;

    Excel.Cells[PhoneRowStart-1,PhoneColumn+1] := 'Клиент';
    Excel.Cells[PhoneRowStart-1,PhoneColumn+2] := 'Заказ';
    Excel.Cells[PhoneRowStart-1,PhoneColumn+3] := 'Склад';
    Excel.Cells[PhoneRowStart-1,PhoneColumn+4] := 'Приемшик';
    Excel.Cells[PhoneRowStart-1,PhoneColumn+5] := 'Сумма заказа';
    Excel.Cells[PhoneRowStart-1,PhoneColumn+6] := 'Новый клиент';


    Try
    Connection := TFDConnection.Create(nil);
      Try
      Connection.DriverName := 'FB';
      with Connection.Params as TFDPhysFBConnectionDefParams do
        begin
          //Protocol := ipTCPIP;
          //Server := 'mail.apetta.ru';
          //Port := 26508;
          Server := '192.168.0.50';
          Database := 'E:\12345\DB\ARM.fdb';
          UserName := 'sysdba';
          Password := 'masterkey';
          IBAdvanced := 'config=WireCompression=false';
        end;
      Connection.Connected := True;
      Except
      on E:Exception do
        begin
          ShowMessage('Произошла ошибка при подключении к БД.'+#13#10+E.ClassName+' '+E.Message+#13#10+'Обратитесь к тех. поддержке!');
          Exit;
        end;
      End;

      Query := TFDQuery.Create(nil);
      Query.ResourceOptions.CmdExecMode := amCancelDialog;
      Query.Connection := Connection;
      For I := PhoneRowStart to PhoneRowEnd do
        begin
          Query.SQL.Clear;
          Phone := VarToStr(Excel.Cells[I,PhoneColumn]);
          Phone := StringReplace(Phone, ' ', '', [rfReplaceAll, rfIgnoreCase]);
          Phone := StringReplace(Phone, '(', '', [rfReplaceAll, rfIgnoreCase]);
          Phone := StringReplace(Phone, ')', '', [rfReplaceAll, rfIgnoreCase]);
          Phone := StringReplace(Phone, '+7', '', [rfReplaceAll, rfIgnoreCase]);
          if (Phone[1]='8') and ((Phone[2]<>'1') and (Phone[3]<>'2'))  then
          Delete(Phone,1,1);
          if Phone[1]='7' then
          Delete(Phone,1,1);
          if Length(Phone)<6 then  continue;
          query.SQL.Text := 'select c.fullname, d.doc_num, s.name, u.description, d.kredit, (case when (c.first_order_doc_id is null) then ''Новый'' when (c.first_order_doc_id=c.last_order_doc_id) then ''Новый'' else ''Старый'' end) as contragnew '+
                            'from contragents c '+
                            'left join docs d on d.doc_id=c.last_order_doc_id '+
                            'left join docs_order dor on dor.doc_id=c.last_order_doc_id '+
                            'left join users u on u.user_id=d.user_id '+
                            'left join sclads s on s.id=dor.sclad_kredit_id '+
                            'where c.teleph_cell like ''%'+Phone+'%'' and c.is_active=1 and c.is_deleted=0';
          Try
            query.Open;
          Except
            Excel.Cells[I,PhoneColumn+1] := '';
            Excel.Cells[I,PhoneColumn+2] := '';
            Excel.Cells[I,PhoneColumn+3] := '';
            Excel.Cells[I,PhoneColumn+4] := '';
            Excel.Cells[I,PhoneColumn+5] := '';
            Excel.Cells[I,PhoneColumn+6] := '';
            Continue;
          End;

          if query.RecordCount>0 then
          begin
            Excel.Cells[I,PhoneColumn+1] := Query.FieldByName('fullname').AsString;
            Excel.Cells[I,PhoneColumn+2] := Query.FieldByName('doc_num').AsString;
            Excel.Cells[I,PhoneColumn+3] := Query.FieldByName('name').AsString;
            Excel.Cells[I,PhoneColumn+4] := Query.FieldByName('description').AsString;
            Excel.Cells[I,PhoneColumn+5] := Query.FieldByName('kredit').AsString;
            Excel.Cells[I,PhoneColumn+6] := Query.FieldByName('contragnew').AsString;
          end else
          begin
            Excel.Cells[I,PhoneColumn+1] := '';
            Excel.Cells[I,PhoneColumn+2] := '';
            Excel.Cells[I,PhoneColumn+3] := '';
            Excel.Cells[I,PhoneColumn+4] := '';
            Excel.Cells[I,PhoneColumn+5] := '';
            Excel.Cells[I,PhoneColumn+6] := '';
          end;
        end;

    Finally
      If Assigned(Query) then FreeAndNil(Query);
      If Assigned(Connection) then FreeAndNil(Connection);
    End;
  finally
    Excel := unassigned;
    Application.Terminate;
  end;
end;

end.
