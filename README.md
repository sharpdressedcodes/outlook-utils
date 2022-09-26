# Outlook Utilities

Work in progress.

## Move sent email to mapped folder
Make a copy of data.example.txt, save it as data.txt. Save it in the following format:

```
someone@example.com,Folder\Path
```

1. At the top of Outlook's `ThisOutlookSession` file, put the correct path to your data file.
2. Import the modules into Outlook VBA.
3. Open local `ThisOutlookSession.cls` and copy over whatever isn't in Outlook's `ThisOutlookSession` file.


`Outlook\Sorted\` will be prepended to the folder path, and `\Sent` will be appended. This can be changed to suit   your folder setup.