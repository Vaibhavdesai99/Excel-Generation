namespace myapp;

entity PipeDetails {
  key id : UUID;
  srno : Integer;
  pipeuniqueid : String(50);
  heatno : String(50);
  coatingno : String(50);
  length : Decimal(10, 2);
  benddetail : String(100);
  weldjointno : String(50);
  fitupinspection : String(100);
  welderid : String(50);
  rootpasswelderno : String(50);
  hotpasswelderno : String(50);
  fillerpasswelderno : String(50);
  cappasswelderno : String(50);
  wpsno : String(50);
  remarks : String(255);
}
