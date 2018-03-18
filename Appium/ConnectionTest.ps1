
#-Begin-----------------------------------------------------------------

  #-Includes------------------------------------------------------------
  $Path = "E:\Projects\Appium"

  [Void][System.Reflection.Assembly]::LoadFrom($Path + "\appium-dotnet-driver.dll")
  [Void][System.Reflection.Assembly]::LoadFrom($Path + "\Castle.Core.dll")
  [Void][System.Reflection.Assembly]::LoadFrom($Path + "\Newtonsoft.Json.dll")
  [Void][System.Reflection.Assembly]::LoadFrom($Path + "\WebDriver.dll")
  [Void][System.Reflection.Assembly]::LoadFrom($Path + "\WebDriver.Support.dll")


  #-Sub Main------------------------------------------------------------
  Function Main() {
  	
    [OpenQA.Selenium.Remote.DesiredCapabilities]$Capabilities = `
      [OpenQA.Selenium.Remote.DesiredCapabilities]::new();
    $Capabilities.SetCapability("deviceName", "android-af34504910a2d1c9");
    $Capabilities.SetCapability("platformVersion", "7.1.2");
    $Capabilities.SetCapability("fullReset", "True");
    $Capabilities.SetCapability([OpenQA.Selenium.Appium.Enums.MobileCapabilityType]::App, "Browser")
    $Capabilities.SetCapability("platformName", "Android");

    [System.Uri]$Uri = [System.Uri]::new("http://127.0.0.1:4723/wd/hub");

    $Driver = `
      [OpenQA.Selenium.Appium.Android.AndroidDriver[OpenQA.Selenium.Appium.AppiumWebElement]]::new($Uri, $Capabilities);

    If ($Driver -eq $null) {
      Return;
    }

    $Driver.Navigate().GoToUrl("https://www.google.de");
  }

  #-Main----------------------------------------------------------------
  Main

#-End-------------------------------------------------------------------
