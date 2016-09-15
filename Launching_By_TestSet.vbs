set tdc = createobject("TDApiOle80.TDConnection")
tdc.InitConnectionEx "https://qualitycenter1152-gbis.fr.world.socgen/qcbin/"
tdc.login "",""
tdc.Connect "RRF","POS_BACARDI"

Set objShell = CreateObject("WScript.Shell")
Set TSetFact = tdc.TestSetFactory
Set tsTreeMgr = tdc.TestSetTreeManager
Set tsFolder = tsTreeMgr.NodeByPath("Root\NightlyRun_New\NIghtlyRun2016\STP-Booking")
Set tsList = tsFolder.FindTestSets("OTC-LST-Booking-MRG")
 
Set theTestSet = tsList.Item(1)
Set Scheduler = theTestSet.StartExecution("")
Scheduler.RunAllLocally = True
Scheduler.run

Set execStatus = Scheduler.ExecutionStatus

Do While RunFinished = False
 execStatus.RefreshExecStatusInfo "all", True
 RunFinished = execStatus.Finished
 Set EventsList = execStatus.EventsList

 For Each ExecEventInfoObj in EventsList
  strNowEvent = ExecEventInfoObj.EventType
 Next

 For i= 1 to execstatus.count
  Set TestExecStatusobj =execstatus.Item(i)
  intTestid = TestExecStatusobj.TestInstance
 Next
Loop
