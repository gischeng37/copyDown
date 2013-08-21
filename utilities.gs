// Some of this code was borrowed and modified from the Flubaroo Script author Dave Abouav
// It anonymously tracks script usage to Google Analytics, allowing our non-profit organization to report the impact of this work to funders
// For original source see http://www.edcode.org

function copyDown_logCopyDown()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Formulas%20Copied%20Down", scriptName, scriptTrackingId, systemName)
}


function copyDown_logFirstInstall()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("First%20Install", scriptName, scriptTrackingId, systemName)
}


function copyDown_logRepeatInstall()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Repeat%20Install", scriptName, scriptTrackingId, systemName)
}


function setCopyDownSid()
{ 
  var copydown_sid = ScriptProperties.getProperty("copydown_sid");
  if (copydown_sid == null || copydown_sid == "")
    {
      // user has never installed formMule before (in any spreadsheet)
      var dt = new Date();
      var ms = dt.getTime();
      var ms_str = ms.toString();
      ScriptProperties.setProperty("copydown_sid", ms_str);
      var copydown_uid = UserProperties.getProperty("copydown_uid");
      if (copydown_uid != null || copydown_uid != "") {
        copyDown_logRepeatInstall();
      }else{
        copyDown_logFirstInstall();
      }
    }
}

