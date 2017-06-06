def myTestCase = context.testCase 
def runner
 
propTestStep = myTestCase.getTestStepByName("VarProperties") // get the Property TestStep
endLoop = propTestStep.getPropertyValue("StopLoop").toString()
 
if (endLoop.toString() == "T" || endLoop.toString()=="True" || endLoop.toString()=="true")
{
 log.info ("Exit Groovy Data Source Looper")
 assert true
}
else
{
 testRunner.gotoStepByName("DataDriver") //setStartStep
}