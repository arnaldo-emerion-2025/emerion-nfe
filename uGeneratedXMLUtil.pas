unit uGeneratedXMLUtil;

interface

uses
  System.Classes, System.Generics.Collections, Xml.XMLIntf;

type
  TgRed = class
    pRedAliq: Double;
    pAliqEfet: Double;
  end;

  TgIBSUF = class
  public
    pIBSUF: Double;
    gRed: TgRed;
    vIBSUF: Double;
  end;

  TgIBSMun = class
    pIBSMun: Double;
    gRed: TgRed;
    vIBSMun: Double;
  end;

  TgCBS = class
    pCBS: Double;
    gRed: TgRed;
    vCBS: Double;
  end;

  TgIBSCBS = class
  public
    vBC: Double;
    gIBSUF: TgIBSUF;
    gIBSMun: TgIBSMun;
    vIBS: Double;
    gCBS: TgCBS;
  end;

  TDetItemNovosTributos = class
  public
    ItemNumber: Integer;
    cst: String;
    cClassTrib: String;
    gIBSCBS: TgIBSCBS;
  end;

  TTotalGIBSUF = class
    vDif: Double;
    vDevTrib: Double;
    vIBSUF: Double;
  end;

  TTotalGIBSMun = class
    vDif: Double;
    vDevTrib: Double;
    vIBSMun: Double;
  end;

  TTotalGIBS = class
    gIBSUF: TTotalGIBSUF;
    gIBSMun: TTotalGIBSMun;
    vIBS: Double;
    vCredPres: Double;
    vCredPresCondSus: Double;
  end;

  TTotalGCBS = class
    vDif: Double;
    vDevTrib: Double;
    vCBS: Double;
    vCredPres: Double;
    vCredPresCondSus: Double;
  end;

  TTotalNovosTributos = class
  public
    vBCIBSCBS: Double;
    gIBS: TTotalGIBS;
    gCBS: TTotalGCBS;

    constructor Create;
  end;

  TInfNFeNovosTributos = class
  public
    Items: TObjectList<TDetItemNovosTributos>;
    Total: TTotalNovosTributos;

    constructor Create;
    destructor Destroy; override;

    procedure LoadFromXML(const XMLText: string);
    function FindByItemNumber(AItemNumber: Integer): TDetItemNovosTributos;
  end;

implementation

uses
  Xml.XMLDoc, System.SysUtils;

{ TTotal }

constructor TTotalNovosTributos.Create;
begin
  inherited;
end;

{ TInfNFe }
constructor TInfNFeNovosTributos.Create;
begin
  inherited;
  Items := TObjectList<TDetItemNovosTributos>.Create(True);
  Total := TTotalNovosTributos.Create;
end;

destructor TInfNFeNovosTributos.Destroy;
begin
  Items.Free;
  Total.Free;
  inherited;
end;

// --------------------------------------------
// Helper to safely get a node child text
// --------------------------------------------
function GetNodeText(Node: IXMLNode; const ChildName: string): string;
var
  C: IXMLNode;
begin
  Result := '';
  if Node = nil then
    Exit;

  C := Node.ChildNodes.FindNode(ChildName);
  if C <> nil then
    Result := C.Text;
end;

function GetNodeDouble(Node: IXMLNode; const ChildName: string): Double;
var
  C: IXMLNode;
begin
  FormatSettings.DecimalSeparator := '.';
  Result := 0;
  if Node = nil then
    Exit;

  C := Node.ChildNodes.FindNode(ChildName);
  if C <> nil then
    Result := StrToFloat(C.Text);
end;

{ TInfNFe.LoadFromXML }

function TInfNFeNovosTributos.FindByItemNumber(AItemNumber: Integer): TDetItemNovosTributos;
var
  Item: TDetItemNovosTributos;
begin
  Result := nil;

  for Item in Items do
    if Item.ItemNumber = AItemNumber then
      Exit(Item);
end;

procedure TInfNFeNovosTributos.LoadFromXML(const XMLText: string);
var
  Doc: IXMLDocument;
  Root, Node, DetNode, ImpNode, IBSCBS, gIBSCBS, gRed: IXMLNode;
  i: Integer;
  Item: TDetItemNovosTributos;
begin
  Doc := TXMLDocument.Create(nil);
  Doc.Options := [doNodeAutoCreate, doNamespaceDecl];
  Doc.ParseOptions := [poPreserveWhiteSpace];
  Doc.Active := True;
  Doc.LoadFromXML(XMLText);

  Root := Doc.DocumentElement;
  if Root = nil then
    Exit;

  // --------------------------
  // READ <det>
  // --------------------------
  for i := 0 to Root.ChildNodes.Count - 1 do
  begin
    DetNode := Root.ChildNodes[i];

    if SameText(DetNode.NodeName, 'det') then
    begin
      Item := TDetItemNovosTributos.Create;
      Item.ItemNumber := StrToIntDef(DetNode.Attributes['nItem'], 0);

      ImpNode := DetNode.ChildNodes.FindNode('imposto');
      if ImpNode <> nil then
      begin
        IBSCBS := ImpNode.ChildNodes.FindNode('IBSCBS');
        if IBSCBS <> nil then
        begin
          Item.cst := GetNodeText(IBSCBS, 'CST');
          Item.cClassTrib := GetNodeText(IBSCBS, 'cClassTrib');

          gIBSCBS := IBSCBS.ChildNodes.FindNode('gIBSCBS');
          if gIBSCBS <> nil then
          begin
            Item.gIBSCBS := TgIBSCBS.Create;
            Item.gIBSCBS.vBC := GetNodeDouble(gIBSCBS, 'vBC');
            Item.gIBSCBS.vIBS := GetNodeDouble(gIBSCBS, 'vIBS');

            Item.gIBSCBS.gIBSUF := TgIBSUF.Create;
            Item.gIBSCBS.gIBSUF.pIBSUF := GetNodeDouble(gIBSCBS.ChildNodes.FindNode('gIBSUF'), 'pIBSUF');
            Item.gIBSCBS.gIBSUF.gRed := TgRed.Create;
            gRed := gIBSCBS.ChildNodes.FindNode('gIBSUF').ChildNodes.FindNode('gRed');
            Item.gIBSCBS.gIBSUF.gRed.pRedAliq := GetNodeDouble(gRed, 'pRedAliq');
            Item.gIBSCBS.gIBSUF.gRed.pAliqEfet := GetNodeDouble(gRed, 'pAliqEfet');
            Item.gIBSCBS.gIBSUF.vIBSUF := GetNodeDouble(gIBSCBS.ChildNodes.FindNode('gIBSUF'), 'vIBSUF');

            Item.gIBSCBS.gIBSMun := TgIBSMun.Create;
            Item.gIBSCBS.gIBSMun.pIBSMun := GetNodeDouble(gIBSCBS.ChildNodes.FindNode('gIBSMun'), 'pIBSMun');
            Item.gIBSCBS.gIBSMun.gRed := TgRed.Create;
            gRed := gIBSCBS.ChildNodes.FindNode('gIBSMun').ChildNodes.FindNode('gRed');
            Item.gIBSCBS.gIBSMun.gRed.pRedAliq := GetNodeDouble(gRed, 'pRedAliq');
            Item.gIBSCBS.gIBSMun.gRed.pAliqEfet := GetNodeDouble(gRed, 'pAliqEfet');
            Item.gIBSCBS.gIBSMun.vIBSMun := GetNodeDouble(gIBSCBS.ChildNodes.FindNode('gIBSMun'), 'vIBSMun');

            Item.gIBSCBS.gCBS := TgCBS.Create;
            Item.gIBSCBS.gCBS.pCBS := GetNodeDouble(gIBSCBS.ChildNodes.FindNode('gCBS'), 'pCBS');
            Item.gIBSCBS.gCBS.gRed := TgRed.Create;
            gRed := gIBSCBS.ChildNodes.FindNode('gCBS').ChildNodes.FindNode('gRed');
            Item.gIBSCBS.gCBS.gRed.pRedAliq := GetNodeDouble(gRed, 'pRedAliq');
            Item.gIBSCBS.gCBS.gRed.pAliqEfet := GetNodeDouble(gRed, 'pAliqEfet');
            Item.gIBSCBS.gCBS.vCBS := GetNodeDouble(gIBSCBS.ChildNodes.FindNode('gCBS'), 'vCBS');
          end;
        end;
      end;

      Items.Add(Item);
    end;
  end;

  // --------------------------
  // READ <total>
  // --------------------------
  Node := Root.ChildNodes.FindNode('total');
  if Node <> nil then
  begin
    Node := Node.ChildNodes.FindNode('IBSCBSTot');
    if Node <> nil then
    begin
      Total.vBCIBSCBS := GetNodeDouble(Node, 'vBCIBSCBS');

      if Node.ChildNodes.FindNode('gIBS') <> nil then
      begin
        Total.gIBS := TTotalGIBS.Create;
        Total.gIBS.gIBSUF := TTotalGIBSUF.Create;
        Total.gIBS.gIBSMun := TTotalGIBSMun.Create;

        Total.gIBS.gIBSUF.vDif := GetNodeDouble(Node.ChildNodes['gIBS'].ChildNodes['gIBSUF'], 'vDif');
        Total.gIBS.gIBSUF.vDevTrib := GetNodeDouble(Node.ChildNodes['gIBS'].ChildNodes['gIBSUF'], 'vDevTrib');
        Total.gIBS.gIBSUF.vIBSUF := GetNodeDouble(Node.ChildNodes['gIBS'].ChildNodes['gIBSUF'], 'vIBSUF');

        Total.gIBS.gIBSMun.vDif := GetNodeDouble(Node.ChildNodes['gIBS'].ChildNodes['gIBSMun'], 'vDif');
        Total.gIBS.gIBSMun.vDevTrib := GetNodeDouble(Node.ChildNodes['gIBS'].ChildNodes['gIBSMun'], 'vDevTrib');
        Total.gIBS.gIBSMun.vIBSMun := GetNodeDouble(Node.ChildNodes['gIBS'].ChildNodes['gIBSMun'], 'vIBSMun');

        Total.gIBS.vIBS := GetNodeDouble(Node.ChildNodes['gIBS'], 'vIBS');
        Total.gIBS.vCredPres := GetNodeDouble(Node.ChildNodes['gIBS'], 'vCredPres');
        Total.gIBS.vCredPresCondSus := GetNodeDouble(Node.ChildNodes['gIBS'], 'vCredPresCondSus');

      end;

      if Node.ChildNodes.FindNode('gCBS') <> nil then
      begin
        Total.gCBS := TTotalGCBS.Create;
        Total.gCBS.vDif := GetNodeDouble(Node.ChildNodes['gCBS'], 'vDif');
        Total.gCBS.vDevTrib := GetNodeDouble(Node.ChildNodes['gCBS'], 'vDevTrib');
        Total.gCBS.vCBS := GetNodeDouble(Node.ChildNodes['gCBS'], 'vCBS');
        Total.gCBS.vCredPres := GetNodeDouble(Node.ChildNodes['gCBS'], 'vCredPres');
        Total.gCBS.vCredPresCondSus := GetNodeDouble(Node.ChildNodes['gCBS'], 'vCredPresCondSus');
      end;
    end;
  end;
end;

end.
