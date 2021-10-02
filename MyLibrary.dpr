    library MyLibrary;

{ Important note about DLL memory management: ShareMem must be the
  first unit in your library's USES clause AND your project's (select
  Project-View Source) USES clause if your DLL exports any procedures or
  functions that pass strings as parameters or function results. This
  applies to all strings passed to and from your DLL--even those that
  are nested in records and classes. ShareMem is the interface unit to
  the BORLNDMM.DLL shared memory manager, which must be deployed along
  with your DLL. To avoid using BORLNDMM.DLL, pass string information
  using PChar or ShortString parameters. }

uses
  System.SysUtils,
  System.Classes,
  System.Generics.Collections,
  System.Variants,
  Comobj;
type
  Product=class
    Private
    ProductCode:string;
    quantity: Double;
    public
    Constructor Create(PC:string;Qu:Double); overload;
  end;
{$R *.res}
{ Product }
Constructor Product.Create(PC: string; Qu: Double);
begin
   ProductCode:=PC;
   quantity:=Qu;
end;
function TotalProduct(var arr:Variant): Variant;stdcall;
var
  dic:TDictionary<string,Product>;
  i: UInt32;
  Key:string;
  Value:Double;
  PD:Product;
  data:Variant;
begin
  dic:=TDictionary<string,Product>.Create;
  try
  for i := VarArrayLowBound(arr,1) to VarArrayHighBound(arr,1) do
    begin
      Key:=TvarData(arr[i,1]).VOleStr;
      Value:=tvardata(arr[i,2]).vdouble;
      if dic.ContainsKey(Key) then
        begin
        PD:=dic.Items[key];
        PD.quantity:=PD.quantity + Value;
        end
      else
        begin
        PD:=Product.Create(Key,Value);
        dic.Add(Key,PD);
        end;
    end;
  data:=VarArrayCreate([1,dic.Count,1,2],varOleStr);
  i:=0;
  for PD in dic.Values do
  begin
    i:=i+1;
    data[i,1]:=PD.ProductCode;
    data[i,2]:=PD.quantity;
  end;
  if PD<>nil then PD.Free;
  finally
    dic.Free;
    Result:=data;
  end;
end;
exports
  TotalProduct;
begin
end.
