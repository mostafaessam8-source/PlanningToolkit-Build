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
                  <button id="importXer" label="Import XER" imageMso="ImportTextFile" size="large" onAction="OnImportXer" screentip="Import every XER table into a new Excel workbook." />
                  <button id="validateXer" label="Validate XER" imageMso="ReviewProtectWorkbook" onAction="OnValidateXer" screentip="Check XER structure, row widths and key identifiers." />
                  <button id="exportXer" label="Export XER" imageMso="ExportExcel" size="large" onAction="OnExportXer" screentip="Export safe worksheet edits to a validated XER with automatic backup and round-trip verification." />
                </group>
                <group id="scheduleGroup" label="Schedule Tools">
                  <button id="formatWbs" label="Build WBS Schedule" imageMso="GroupRows" size="large" onAction="OnBuildSchedule" screentip="Create a grouped WBS schedule with summary rows and critical path." />
                  <button id="createGantt" label="Create Gantt Chart" imageMso="ChartTypeBarInsertGallery" size="large" onAction="OnBuildSchedule" screentip="Create the WBS schedule and weekly Gantt chart from an XER file." />
                  <button id="compareXer" label="Compare Baseline / Update" imageMso="ReviewCompareTwoVersions" size="large" onAction="OnCompareXer" screentip="Compare a baseline XER against an updated XER." />
                  <button id="createPms" label="Create PMS Dashboard" imageMso="PivotTableInsert" onAction="OnCreatePms" screentip="Create dashboard, S-curve, look-ahead and critical path from baseline and update XER files." />
                  <button id="exportReportPdf" label="Export Report as PDF" imageMso="FileSaveAsPdfOrXps" onAction="OnExportReportPdf" screentip="Export the active Planning Toolkit report workbook to a single branded, print-ready PDF." />
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
    public void OnImportXer(IRibbonControl control) => RibbonActions.ImportXer();
    public void OnExportXer(IRibbonControl control) => RibbonActions.ExportXer();
    public void OnValidateXer(IRibbonControl control) => RibbonActions.ValidateXer();
    public void OnBuildSchedule(IRibbonControl control) => RibbonActions.BuildSchedule();
    public void OnCompareXer(IRibbonControl control) => RibbonActions.CompareXer();
    public void OnCreatePms(IRibbonControl control) => RibbonActions.CreatePms();
    public void OnExportReportPdf(IRibbonControl control) => RibbonActions.ExportReportPdf();
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
