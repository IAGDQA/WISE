using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Selenium;
using System.Diagnostics;

namespace AdvSeleniumAPI
{
    public class AdvSeleniumAPIv2 : ISelenium
    {
        System.Collections.ArrayList BaseListOfResItems;//result items
        ResultClass res;

        ISelenium selenium;

        public AdvSeleniumAPIv2()
        {
            BaseListOfResItems = new System.Collections.ArrayList();
            
        }

        public void StartupServer(string _ip)
        {
            selenium =
                new DefaultSelenium("localhost", 4444, "*firefox", _ip);//firefox//googlechrome//iexplore
            selenium.Start();
            selenium.SetSpeed("500");
            selenium.Open("/config");
            selenium.WindowMaximize();
            System.Threading.Thread.Sleep(1000);

        }
        public void StartupServer(string _ip, string _bw)
        {
            if(_bw == "Chrome")
            {
                selenium =
                new DefaultSelenium("localhost", 4444, "*googlechrome", _ip);//firefox//googlechrome//iexplore
                selenium.Start();
                selenium.SetSpeed("500");
                selenium.Open("/config");
                selenium.WindowMaximize();
                System.Threading.Thread.Sleep(1000);
            }
            else
            {
                selenium =
                new DefaultSelenium("localhost", 4444, "*iexplore", _ip);//firefox//googlechrome//iexplore
                selenium.Start();
                selenium.SetSpeed("500");
                selenium.Open("/config");
                selenium.WindowMaximize();
                System.Threading.Thread.Sleep(1000);
            }

        }

        public System.Collections.ArrayList GetStepResult()
        {
            var temp = BaseListOfResItems.Clone();
            BaseListOfResItems.Clear();
            return (System.Collections.ArrayList)temp;
        }

        //===============================================================//
        
        

        ///////////////////////////////////////////////////////////////////////////////////////////////

        public void AddLocationStrategy(string strategyName, string functionDefinition) { }

        public void AddScript(string scriptContent, string scriptTagId) { }

        public void AddSelection(string locator, string optionLocator) { }

        public void AllowNativeXpath(string allow) { }

        public void AltKeyDown() { }

        public void AltKeyUp() { }

        public void AnswerOnNextPrompt(string answer) { }

        public void AssignId(string locator, string identifier) { }

        public void AttachFile(string fieldLocator, string fileLocator) { }

        public void CaptureEntirePageScreenshot(string filename, string kwargs) { }

        public string CaptureEntirePageScreenshotToString(string kwargs) { return ""; }

        public void CaptureScreenshot(string filename) { }

        public string CaptureScreenshotToString() { return ""; }

        public void Check(string locator) { }

        public void ChooseCancelOnNextConfirmation() { }

        public void ChooseOkOnNextConfirmation() { }

        public void Click(string locator)
        {
            bool actRes = false;string errStr = "";
            try
            {
                selenium.Click(locator);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "Click, " + locator,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);
        }

        public void ClickAt(string locator, string coordString) { }

        public void Close()
        {            
            selenium.Close();
        }

        public void ContextMenu(string locator) { }

        public void ContextMenuAt(string locator, string coordString) { }

        public void ControlKeyDown() { }

        public void ControlKeyUp() { }

        public void CreateCookie(string nameValuePair, string optionsString) { }

        public void DeleteAllVisibleCookies() { }

        public void DeleteCookie(string name, string optionsString) { }

        public void DeselectPopUp()
        {
            bool actRes = false; string errStr = "";
            try
            {
                selenium.DeselectPopUp();
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "DeselectPopUp",
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);
        }

        public void DoubleClick(string locator) { }

        public void DoubleClickAt(string locator, string coordString) { }

        public void DragAndDrop(string locator, string movementsString) { }

        public void DragAndDropToObject(string locatorOfObjectToBeDragged, string locatorOfDragDestinationObject) { }

        public void Dragdrop(string locator, string movementsString) { }

        public void FireEvent(string locator, string eventName) { }

        public void Focus(string locator) { }

        public string GetAlert() { return ""; }

        public string[] GetAllButtons() { return null; }

        public string[] GetAllFields() { return null; }

        public string[] GetAllLinks() { return null; }

        public string[] GetAllWindowIds()
        {
            bool actRes = false; string errStr = "";
            string[] resStrArray = new string[1];
            try
            {
                resStrArray = selenium.GetAllWindowIds();
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "GetAllWindowIds",
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);

            return resStrArray;
        }

        public string[] GetAllWindowNames()
        {
            bool actRes = false; string errStr = "";
            string[] resStrArray = new string[1];
            try
            {
                resStrArray = selenium.GetAllWindowNames();
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "GetAllWindowNames",
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);

            return resStrArray;
        }

        public string[] GetAllWindowTitles() { return null; }

        public string GetAttribute(string attributeLocator)
        {
            bool actRes = false; string errStr = "", resStr = "";
            try
            {
                resStr = selenium.GetAttribute(attributeLocator);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "GetAttribute, " + attributeLocator,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);

            return resStr;
        }

        public string[] GetAttributeFromAllWindows(string attributeName) { return null; }

        public string GetBodyText() { return null; }

        public string GetConfirmation() { return null; }

        public string GetCookie() { return null; }

        public string GetCookieByName(string name) { return null; }

        public decimal GetCSSCount(string cssLocator) { return 0; }

        public decimal GetCursorPosition(string locator) { return 0; }

        public decimal GetElementHeight(string locator) { return 0; }

        public decimal GetElementIndex(string locator) { return 0; }

        public decimal GetElementPositionLeft(string locator) { return 0; }

        public decimal GetElementPositionTop(string locator) { return 0; }

        public decimal GetElementWidth(string locator) { return 0; }

        public string GetEval(string script) { return null; }

        public string GetExpression(string expression) { return null; }

        public string GetHtmlSource() { return null; }

        public string GetLocation() { return null; }

        public decimal GetMouseSpeed() { return 0; }

        public string GetPrompt() { return null; }

        public string GetSelectedId(string selectLocator) { return null; }

        public string[] GetSelectedIds(string selectLocator) { return null; }

        public string GetSelectedIndex(string selectLocator)
        {
            bool actRes = false; string resStr = "", errStr = "";
            try
            {
                resStr = selenium.GetSelectedIndex(selectLocator);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "GetSelectedIndex, " + selectLocator,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);

            return resStr;
        }

        public string[] GetSelectedIndexes(string selectLocator) { return null; }

        public string GetSelectedLabel(string selectLocator)
        {
            bool actRes = false; string resStr = "", errStr = "";
            try
            {
                resStr = selenium.GetSelectedLabel(selectLocator);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "GetSelectedLabel, " + selectLocator,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);

            return resStr;
        }

        public string[] GetSelectedLabels(string selectLocator) { return null; }

        public string GetSelectedValue(string selectLocator) { return null; }

        public string[] GetSelectedValues(string selectLocator) { return null; }

        public string[] GetSelectOptions(string selectLocator) { return null; }

        public string GetSpeed() { return null; }

        public string GetTable(string tableCellAddress) { return null; }

        public string GetText(string locator)
        {
            bool actRes = false; string resStr = "", errStr = "";
            try
            {
                resStr = selenium.GetText(locator);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "GetText, " + locator,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);

            return resStr;
        }

        public string GetTitle()
        {
            bool actRes = false; string errStr = "", resStr = "";
            try
            {
                resStr = selenium.GetTitle();
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "GetValue",
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);

            return resStr;
        }

        public string GetValue(string locator)
        {
            bool actRes = false; string errStr = "", resStr = "";
            try
            {
                resStr = selenium.GetValue(locator);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "GetValue, " + locator,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);

            return resStr;

        }

        public bool GetWhetherThisFrameMatchFrameExpression(string currentFrameString, string target) { return false; }

        public bool GetWhetherThisWindowMatchWindowExpression(string currentWindowString, string target) { return false; }

        public decimal GetXpathCount(string xpath) { return 0; }

        public void GoBack() { }

        public void Highlight(string locator) { }

        public void IgnoreAttributesWithoutValue(string ignore) { }

        public bool IsAlertPresent() { return false; }

        public bool IsChecked(string locator) { return false; }

        public bool IsConfirmationPresent() { return false; }

        public bool IsCookiePresent(string name) { return false; }

        public bool IsEditable(string locator) { return false; }

        public bool IsElementPresent(string locator) { return false; }

        public bool IsOrdered(string locator1, string locator2) { return false; }

        public bool IsPromptPresent() { return false; }

        public bool IsSomethingSelected(string selectLocator) { return false; }

        public bool IsTextPresent(string pattern) { return false; }

        public bool IsVisible(string locator) { return false; }

        public void KeyDown(string locator, string keySequence) { }

        public void KeyDownNative(string keycode) { }

        public void KeyPress(string locator, string keySequence) { }

        public void KeyPressNative(string keycode) { }

        public void KeyUp(string locator, string keySequence) { }

        public void KeyUpNative(string keycode) { }

        public void MetaKeyDown() { }

        public void MetaKeyUp() { }

        public void MouseDown(string locator)
        {
            selenium.MouseDown(locator);
        }

        public void MouseDownAt(string locator, string coordString) { }

        public void MouseDownRight(string locator) { }

        public void MouseDownRightAt(string locator, string coordString) { }

        public void MouseMove(string locator) { }

        public void MouseMoveAt(string locator, string coordString) { }

        public void MouseOut(string locator) { }

        public void MouseOver(string locator) { }

        public void MouseUp(string locator)
        {
            selenium.MouseUp(locator);
        }

        public void MouseUpAt(string locator, string coordString) { }

        public void MouseUpRight(string locator) { }

        public void MouseUpRightAt(string locator, string coordString) { }

        public void Open(string url) { selenium.Open(url); }

        public void OpenWindow(string url, string windowID) { }

        public void Refresh() { }

        public void RemoveAllSelections(string locator) { }

        public void RemoveScript(string scriptTagId) { }

        public void RemoveSelection(string locator, string optionLocator) { }

        public string RetrieveLastRemoteControlLogs() { return null; }

        public void Rollup(string rollupName, string kwargs) { }

        public void RunScript(string script) { }

        public void Select(string selectLocator, string optionLocator)
        {
            bool actRes = false; string errStr = "";
            try
            {
                selenium.Select(selectLocator, optionLocator);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "Select, " + selectLocator,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);
        }

        public void SelectFrame(string locator)
        {
            bool actRes = false; string errStr = "";
            try
            {
                selenium.SelectFrame(locator);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "SelectFrame, " + locator,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);
        }

        public void SelectPopUp(string windowID)
        {
            bool actRes = false; string errStr = "";
            try
            {
                selenium.SelectPopUp(windowID);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "SelectPopUp, " + windowID,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);
        }

        public void SelectWindow(string windowID)
        {
            bool actRes = false; string errStr = "";
            try
            {
                selenium.SelectWindow(windowID);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "SelectWindow, " + windowID,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);
        }

        public void SetBrowserLogLevel(string logLevel) { }

        public void SetContext(string context) { }

        public void SetCursorPosition(string locator, string position) { }

        public void SetExtensionJs(string extensionJs) { }

        public void SetMouseSpeed(string pixels) { }

        public void SetSpeed(string value) { selenium.SetSpeed(value); }

        public void SetTimeout(string timeout) { }

        public void ShiftKeyDown() { }

        public void ShiftKeyUp() { }

        public void ShutDownSeleniumServer() { }

        public void Start() { selenium.Start(); }

        public void Stop() { selenium.Stop(); }

        public void Submit(string formLocator)
        {
            bool actRes = false; string errStr = "";
            try
            {
                selenium.Submit(formLocator);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "Submit, " + formLocator,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);
        }

        public void Type(string locator, string value)
        {
            bool actRes = false; string errStr = "";
            try
            {
                selenium.Type(locator, value);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "Type, " + locator,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);
        }

        public void TypeKeys(string locator, string value) { }

        public void Uncheck(string locator) { }

        public void UseXpathLibrary(string libraryName) { }

        public void WaitForCondition(string script, string timeout) { }

        public void WaitForFrameToLoad(string frameAddress, string timeout) { }

        public void WaitForPageToLoad(string timeout) { selenium.WaitForPageToLoad(timeout); }

        public void WaitForPopUp(string windowID, string timeout)
        {
            bool actRes = false; string errStr = "";
            try
            {
                selenium.WaitForPopUp(windowID, timeout);
                actRes = true;
            }
            catch (System.Exception ex)
            {
                errStr = ex.ToString();
            }

            res = new ResultClass()
            {
                Decp = "WaitForPopUp, " + windowID,
                Res = actRes ? "pass" : "fail",
                Tdev = " ms",
                Err = errStr,
            };
            BaseListOfResItems.Add(res);
        }
        public void WindowFocus() { }

        public void WindowMaximize() { selenium.WindowMaximize(); }


    }//class

    public class ResultClass//Result class
    {
        public string Decp//Description
        {
            get;
            set;
        }
        public string Res//Result
        {
            get;
            set;
        }
        public string Tdev//time devision
        {
            get;
            set;
        }
        public string Err//Error code
        {
            get;
            set;
        }
    }

    public enum BrowserType
    {
        FireFox,
        Chrome,
        IE,
    }
}
