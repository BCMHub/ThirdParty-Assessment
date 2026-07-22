# Build on a locked-down corporate laptop

No PowerShell, internet connection, or software installation is needed.

1. Copy this whole project folder to the laptop. Keep the `src` folder and all its files together.
2. In Excel, create a blank workbook and save it as a macro-enabled workbook, for example `Tracker_Builder.xlsm`, in the project folder directly beside `src`. Do **not** name it `ThirdParty_Assessment_Tracker.xlsm`.
3. Press **Alt+F11**. In the VBA editor, click **Insert > Module**.
4. Open `Import_Bootstrap.bas` in a text editor, copy all of it, and paste it into the new module. If the pasted module is not named `Import_Bootstrap`, that is okay.
5. In Excel, enable VBA project access: **File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model**. If you changed this setting, close every Excel window, reopen `Tracker_Builder.xlsm`, and enable macros.
6. Press **Alt+F8**, select **BuildTracker**, and click **Run**. Answer the three prompts for the initial Manager, display name, and internal email domain.
7. Wait for the PASS or FAIL message. In the same folder, find:
   - `ThirdParty_Assessment_Tracker.xlsm`
   - `BuildLog_YYYYMMDD_HHNNSS.txt`
8. If the result is FAIL, or anything looks wrong, copy and paste the **entire BuildLog text file** back here. The same log is also available on the `BuildLog` sheet in `Tracker_Builder.xlsm`.

The builder reads only local files from `src`. It does not use PowerShell, make network calls, or create compiled Office artifacts.
