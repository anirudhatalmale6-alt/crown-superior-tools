# Crown Superior - Excel tool modules

The VBA modules for the Crown Superior automation workbook. Download the
file you need, then in Excel:

1. **Alt + F11** to open the VBA editor
2. Right-click the old module of the same name in the left-hand tree,
   **Remove**, and say **No** when it offers to export it
3. **File → Import File**, pick the `.bas` you downloaded
4. Back in Excel, run `CrownVersion` to check which build you are on

To download one file on its own, click it in the list above, then press
the **Download raw file** button at the top right of the code.

| Module | What it does |
|---|---|
| `CrownPayments.bas` | Reads what is due from the carriers - United Auto, Trisura (Verve) - and writes it back to the website |
| `CrownQuotes.bas` | Brings quote requests into the tool and hands one to the Edit Data page |
| `CrownAPI.bas` | The connection to crownsuperior.com that the other two use |
| `CrownUnblock.bas` | Getting past the alerts and pages that stop a United quote |
| `CrownSheet.gs` | Goes on Google, not in Excel - see the top of the file |

No passwords are in these files. The carrier logins stay on the **Input**
sheet of your workbook, and the website key is stored inside the workbook
itself - see `CrownSetKey` in `CrownAPI.bas`.
