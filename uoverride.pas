unit uOverride;

{$mode objfpc}{$H+}
{Dasar-dasar Override. Ditulis dari Contoh Kulgram Mas Kofa https://t.me/kppdi
 Kontributor : Didi Kurniadi}

interface

uses
  Classes, SysUtils, Grids, DB, Zdataset, ZConnection, fpspreadsheet, fpspreadsheetctrls,
  fpDBExport, fpsimplejsonexport, dialogs;

type
  TStringGrid= Class (Grids.TStringGrid)
    Private
    //private declare
    function GetSQL :String;
    procedure SetSQL (ASQL :String);
    function GetConnection :TZConnection;
    procedure SetConnection (Connection:TZConnection);
    function GetDataSet : TZQuery;
    procedure SetDataSet (ADataset : TZQuery);
    procedure CreateDataSet;

    Public
    //public declare
    PSQL, Error : string;
    Conn        : TZConnection;
    DBSource    : TZQuery;
    property Connection : TZConnection Read GetConnection Write SetConnection; //harus pake Zeos
    property SQLText: String Read GetSQL Write SetSQL;  //basic load DB ke StringGrid
    property Dataset : TZQuery Read GetDataset Write SetDataSet; ////Barang kali kapan2 butuh
    procedure RefreshSQL; //re-open database
    procedure UpdateData; // update data dari StringGrid
    procedure DeleteData; //Delete selected data on grid
    procedure SaveToExcel (AFile : string); //save to excel with FPSpreadSheet lib help
    procedure SaveToJSON (AFile : string); //save to JSON format
    procedure SaveToPDF (AFile : string); //save to PDF file

  end;

implementation
procedure TStringGrid.CreateDataSet;
begin
   DBSource := TZQuery.Create(self);
     with DBSource do
            begin
              connection := conn;
              sql.Text:=PSQL;
              open;
            end;
end;

procedure TStringGrid.SaveToPDF (AFile : string);
begin

end;

function TStringGrid.GetDataset : TZQuery;
begin
  result := DBSource;
end;

procedure TStringGrid.SetDataset (ADataset : TZQuery);
begin
  DBSource := TZQuery.Create(self);
  ADataset := DBSource;
end;

procedure TStringGrid.SaveToJSON(AFile :String);
var
  JSONEx : TSimpleJSONExporter;
  iii: integer;
begin
  try
     JSONEx := TSimpleJSONExporter.Create(self);
     CreateDataSet;
     JSONEx.Dataset:=DBSource;

         for iii := 0 to DBSource.FieldCount -1 do
             begin
                JSONEx.ExportFields.AddField(DBSource.Fields[iii].FieldName);
                JSONEx.ExportFields[iii].FieldName:=DBSource.Fields[iii].FieldName;
             end;
     JSONEx.FileName:=AFile;
     JSONEx.Execute;
     RefreshSQL;
   except
      on E: Exception do Error:= 'An exception was raised: ' + E.Message;
    end;
end;

procedure TStringGrid.UpdateData;
var
  ii,jj : integer;
  sss : TStringList;
begin
    try
    sss := TStringList.Create;
    DBsource := TZQuery.Create(self);
    CreateDataSet;
       for ii := 1 to RowCount do
           begin
            for jj := 0 to ColCount-1 do
                begin
                    sss.Add(cells[jj,ii]);
                end;
                    DBSource.Insert;
                    DBSource.Fields[jj].Value :=sss[jj];
                    DBSource.Post;
                    sss.Clear;
           end;
       RefreshSQL;
     except
      on E: Exception do Error:= 'An exception was raised: ' + E.Message;
    end;
end;

procedure TStringGrid.DeleteData;
begin
  try
    DBsource := TZQuery.Create(self);
    CreateDataSet;
    DBsource.Locate(cells[col,0],cells[col,row],[]);
    if not DBSource.IsEmpty then DBsource.Delete;
  finally
    RefreshSQL;
  end;
end;

procedure TStringGrid.RefreshSQL;
begin
    SetSQL(PSQL);
end;

function TStringGrid.GetSQL :String;
begin
  result := PSQL;
end;

procedure TStringGrid.SetConnection (Connection:TZConnection);
begin
  Conn := TZConnection.Create(self);
  Conn := Connection;
end;

function TStringGrid.GetConnection : TZConnection;
begin
  result := Conn;
end;

procedure TStringGrid.SetSQL (ASQL :String);
var
  i,j:integer;
begin
  try
  DBSource := TZQuery.Create(self);

  with DBSource do
       begin
         connection := conn;
         DBSource.SQL.Text:=ASQL;
         Open;
       end;
      PSQL:=ASQL;
      RowCount :=DBSource.RecordCount+1;
      ColCount:=DBSource.FieldCount;

      for i := 0 to DBSource.Fields.Count - 1 do
               cells[i, 0]:= DBSource.Fields[i].FieldName;
      DBSource.First;
      j := 0;
      while not DBSource.EOF do
      begin
        for i := 0 to DBSource.Fields.Count - 1 do
               cells[i, j + 1]:= DBSource.Fields[i].AsString;
        DBSource.Next;
        Inc(j);
      end;
        DBSource.Free;
   except
      on E: Exception do Error:= 'An exception was raised: ' + E.Message;
   end;

end;

procedure TStringGrid.SaveToExcel(AFile:String);
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  i, j: Integer;
begin
  try
      begin
      // Create the spreadsheet
          MyWorkbook := TsWorkbook.Create;
          MyWorksheet := MyWorkbook.AddWorksheet('Worksheet1');

      // Write all cells to the worksheet
            for j:=0 to rowcount - 1 do
            begin
              for i := 0 to ColCount - 1 do
              MyWorksheet.WriteText(j + 1, i+1, cells[i,j]);
            end;

      // Save the spreadsheet to a file
            MyWorkbook.WriteToFile(AFile+'.'+'xlsx');
            MyWorkbook.Free;
      end;
  except
     on E: Exception do Error:= 'An exception was raised: ' + E.Message;
  end;

end;

end.

