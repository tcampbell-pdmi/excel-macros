## Running a Macro in Excel (Windows)

There are a few ways to do it depending on your preference:

---

### Option 1: Via the Developer Tab
1. Open Excel and click the **Developer** tab in the ribbon.
2. Click **Macros** (or press **Alt + F8**).
3. In the dialog box, select the macro you want to run from the list.
4. Click **Run**.

> 💡 **Don't see the Developer tab?** Go to **File → Options → Customize Ribbon**, then check the **Developer** box and click **OK**.

---

### Option 2: Keyboard Shortcut (Fastest)
If a shortcut was assigned to the macro when it was created, just press it directly — e.g., **Ctrl + Shift + M**.

---

### Option 3: From the View Tab
1. Go to **View** in the ribbon.
2. Click **Macros → View Macros**.
3. Select your macro and click **Run**.

---

### Option 4: Via the VBA Editor
1. Press **Alt + F11** to open the Visual Basic Editor.
2. Find your macro in the left-hand project panel.
3. Click inside the macro code and press **F5** to run it.

---


## Updating Testing Env

When the time comes, you can change the target testing env by opening VBA Editor and updating the `API_URL` variable

QA = mirthtest20:10900
Eval = mirth-eval:10900