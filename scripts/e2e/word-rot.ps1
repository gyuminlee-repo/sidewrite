# A document that is open in a running Word, found in the Running Object
# Table by its full path. Dot-sourced by word-own.ps1, word-com.ps1 and
# replay-save.ps1.
#
# Marshal.BindToMoniker is not used for this: a file moniker whose file is
# not (yet) in the table is bound by starting the server of its file type
# (IMoniker::BindToObject of a file moniker: GetClassFile, CoCreateInstance,
# IPersistFile::Load). For a .docx that is a new WINWORD.EXE that no run has
# started and none quits; a lookup that races a Word still loading its
# document starts one. [SwRot]::Find only reads the table: it returns $null
# when the document is not there and never starts anything.
if (-not ("SwRot" -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class SwRot {
  [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable rot);
  [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved, out IBindCtx ctx);
  public static object Find(string path) {
    IRunningObjectTable rot;
    IBindCtx ctx;
    if (GetRunningObjectTable(0, out rot) != 0 || CreateBindCtx(0, out ctx) != 0) return null;
    IEnumMoniker list;
    rot.EnumRunning(out list);
    IMoniker[] one = new IMoniker[1];
    while (list.Next(1, one, IntPtr.Zero) == 0) {
      string name = null;
      try { one[0].GetDisplayName(ctx, null, out name); } catch { name = null; }
      if (name != null && String.Equals(name, path, StringComparison.OrdinalIgnoreCase)) {
        object found;
        if (rot.GetObject(one[0], out found) == 0) return found;
      }
    }
    return null;
  }
}
'@
}
