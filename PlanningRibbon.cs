using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace PlanningToolkit.Excel;

[ComVisible(true)]
public sealed class PlanningRibbon : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId) =>
        """
        <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
          <ribbon>
            <tabs>
              <tab id="planningToolkitTab" label="Planning Toolkit">
                <group id="xerGroup" label="XER Operations">
                  <button id="importXer" label="Import XER (Phase 2)" imageMso="ImportTextFile" size="large" enabled="false" />
                  <button id="validateXer" label="Validate XER (Phase 2)" imageMso="ReviewProtectWorkbook" enabled="false" />
                  <button id="exportXer" label="Export XER (Phase 6)" imageMso="ExportExcel" enabled="false" />
                </group>
                <group id="scheduleGroup" label="Schedule Tools">
                  <button id="formatWbs" label="WBS Tools (Phase 3)" imageMso="GroupRows" size="large" enabled="false" />
                  <button id="createGantt" label="Gantt Chart (Phase 3)" imageMso="ChartTypeBarInsertGallery" size="large" enabled="false" />
                  <button id="compareXer" label="Compare XER (Phase 4)" imageMso="ReviewCompareTwoVersions" size="large" enabled="false" />
                  <button id="createPms" label="Create PMS (Phase 5)" imageMso="PivotTableInsert" enabled="false" />
                </group>
                <group id="textGroup" label="Text &amp; Data">
                  <button id="fillDown" label="Fill Down" imageMso="FillDown" onAction="OnFillDown" screentip="Fill the first selected row down through the selection." />
                  <button id="trimSpaces" label="Trim" imageMso="ClearFormatting" onAction="OnTrimSpaces" screentip="Trim leading, trailing and repeated spaces." />
                  <button id="cleanText" label="Clean" imageMso="Clear" onAction="OnCleanText" screentip="Remove non-printing control characters." />
                  <menu id="changeCase" label="Change Case" imageMso="ChangeCaseMenu">
                    <button id="upperCase" label="UPPER CASE" onAction="OnUpperCase" />
                    <button id="lowerCase" label="lower case" onAction="OnLowerCase" />
                    <button id="properCase" label="Proper Case" onAction="OnProperCase" />
                    <button id="sentenceCase" label="Sentence case" onAction="OnSentenceCase" />
                  </menu>
                  <button id="textToDate" label="Text to Date" imageMso="DateAndTimeInsert" onAction="OnTextToDate" />
                </group>
                <group id="rangeGroup" label="Range Utilities">
                  <button id="splitText" label="Split Text" imageMso="TextToColumns" onAction="OnSplitText" />
                  <button id="mergeText" label="Merge Text" imageMso="MergeAndCenter" onAction="OnMergeText" />
                  <button id="uniqueValues" label="Unique Values" imageMso="RemoveDuplicates" onAction="OnUniqueValues" />
                  <button id="removeBlankRows" label="Remove Blank Rows" imageMso="DeleteRows" onAction="OnRemoveBlankRows" />
                </group>
                <group id="settingsGroup" label="Settings &amp; Help">
                  <button id="settings" label="Settings" imageMso="FileProperties" size="large" onAction="OnSettings" />
                  <button id="viewLogs" label="View Logs" imageMso="FileOpen" onAction="OnViewLogs" />
                  <button id="manual" label="User Guide" imageMso="Help" onAction="OnUserGuide" />
                  <button id="about" label="About" imageMso="Info" onAction="OnAbout" />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>
        """;

    public void OnFillDown(IRibbonControl control) => RibbonActions.FillDown();
    public void OnTrimSpaces(IRibbonControl control) => RibbonActions.TrimSpaces();
    public void OnCleanText(IRibbonControl control) => RibbonActions.CleanText();
    public void OnUpperCase(IRibbonControl control) => RibbonActions.UpperCase();
    public void OnLowerCase(IRibbonControl control) => RibbonActions.LowerCase();
    public void OnProperCase(IRibbonControl control) => RibbonActions.ProperCase();
    public void OnSentenceCase(IRibbonControl control) => RibbonActions.SentenceCase();
    public void OnTextToDate(IRibbonControl control) => RibbonActions.TextToDate();
    public void OnSplitText(IRibbonControl control) => RibbonActions.SplitText();
    public void OnMergeText(IRibbonControl control) => RibbonActions.MergeText();
    public void OnUniqueValues(IRibbonControl control) => RibbonActions.UniqueValues();
    public void OnRemoveBlankRows(IRibbonControl control) => RibbonActions.RemoveBlankRows();
    public void OnSettings(IRibbonControl control) => RibbonActions.ShowSettings();
    public void OnViewLogs(IRibbonControl control) => RibbonActions.ViewLogs();
    public void OnUserGuide(IRibbonControl control) => RibbonActions.OpenUserGuide();
    public void OnAbout(IRibbonControl control) => RibbonActions.ShowAbout();
}
