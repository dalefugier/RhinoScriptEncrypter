using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Rhino;
using Rhino.Commands;
using Rhino.Geometry;
using Rhino.Input;
using Rhino.Input.Custom;

namespace RhinoScriptEncrypter
{
  /// <summary>
  /// RhinoScriptEncrypterCommand
  /// </summary>
  [System.Runtime.InteropServices.Guid("1da690e6-7e5c-4949-9c03-b9cafb333144")]
  public class RhinoScriptEncrypterCommand : Command
  {
    int _optionIndex = 0;

    /// <summary>
    /// Public constructor
    /// </summary>
    public RhinoScriptEncrypterCommand()
    {
    }

    /// <summary>
    /// Returns the English command name
    /// </summary>
    public override string EnglishName
    {
      get { return "RhinoScriptEncrypter"; }
    }

    /// <summary>
    /// RunCommand override
    /// </summary>
    protected override Result RunCommand(RhinoDoc doc, RunMode mode)
    {
      string rhinoScriptFile = string.Empty;

      // Prompt for a filename
      if (mode == RunMode.Interactive)
      {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Filter = "RhinoScript Files (*.rvb)|*.rvb";
        openFileDialog.Title = "Select RhinoScript File";
        if (openFileDialog.ShowDialog() != DialogResult.OK)
          return Result.Cancel;

        rhinoScriptFile = openFileDialog.FileName;
      }
      else
      {
        Result rc = GetOpenFileName("RhinoScript file to encrypt", ref rhinoScriptFile);
        if (rc != Result.Success)
          return rc;
      }

      // Verify the filename
      rhinoScriptFile = rhinoScriptFile.Trim();
      if (string.IsNullOrEmpty(rhinoScriptFile))
        return Result.Nothing;

      // Verify the file exists
      if (!File.Exists(rhinoScriptFile))
      {
        string error = string.Format("RhinoScript file not found - {0}\n", rhinoScriptFile);
        if (mode == RunMode.Interactive)
          MessageBox.Show(error, EnglishName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        else
          RhinoApp.WriteLine(error);
        return Result.Failure;
      }

      // Generate an encryption password
      string encryptPassword = Guid.NewGuid().ToString();

      // Do the file encryption
      string encryptedFileName = string.Empty;
      try
      {
        // Read the script
        string clearString = File.ReadAllText(rhinoScriptFile);
        // Encrypt the script
        string encryptedString = Encrypt(clearString, encryptPassword);
        // Write the encrypted script
        encryptedFileName = Path.ChangeExtension(rhinoScriptFile, ".rvbx");
        File.WriteAllText(encryptedFileName, encryptedString);
      }
      catch (Exception ex)
      {
        string error = string.Format("Error encrypting RhinoScript file - {0}\n{1}\n", rhinoScriptFile, ex.Message);
        if (mode == RunMode.Interactive)
          MessageBox.Show(error, EnglishName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        else
          RhinoApp.WriteLine(error);
        return Result.Failure;
      }

      // Report the results
      string outputString = "RhinoScript file encryption successful!\n";
      outputString += string.Format("  Input filename: {0}\n", rhinoScriptFile);
      outputString += string.Format("  Output filename: {0}\n", encryptedFileName);
      outputString += string.Format("  Decrypt password: {0}\n", encryptPassword);

      if (mode == RunMode.Interactive)
        Rhino.UI.Dialogs.ShowTextDialog(outputString, EnglishName);
      else
      {
        string[] values = new string[] { "HistoryWindow", "File", "Clipboard", "Dialog" };

        GetOption go = new GetOption();
        go.SetCommandPrompt(string.Format("Text destination <{0}>", values[_optionIndex]));
        go.AcceptNothing(true);
        go.AddOption(new Rhino.UI.LocalizeStringPair(values[0], values[0])); // 1
        go.AddOption(new Rhino.UI.LocalizeStringPair(values[1], values[1])); // 2
        go.AddOption(new Rhino.UI.LocalizeStringPair(values[2], values[2])); // 3
        go.AddOption(new Rhino.UI.LocalizeStringPair(values[3], values[3])); // 4
        
        GetResult res = go.Get();

        if (res == GetResult.Option)
          _optionIndex = go.OptionIndex() - 1;
        else if (res != GetResult.Nothing)
          return Result.Cancel;

        switch (_optionIndex)
        {
          case 0: // HistoryWindow
            RhinoApp.WriteLine(outputString);
            break;

          case 1: // File
            {
              string outputFileName = string.Empty;

              Result cmd_rc = GetSaveFileName("Save file name", ref outputFileName);
              if (cmd_rc != Result.Success)
                return Result.Cancel;

              outputFileName = outputFileName.Trim();
              if (string.IsNullOrEmpty(outputFileName))
                return Result.Nothing;

              try
              {
                outputString = outputString.Replace("\n", "\r\n");
                using (StreamWriter stream = new StreamWriter(outputFileName))
                  stream.Write(outputString);
              }
              catch
              {
                RhinoApp.WriteLine("Unable to write to file.\n");
                return Result.Failure;
              }
            }
            break;

          case 2: // Clipboard
            outputString = outputString.Replace("\n", "\r\n");
            Clipboard.SetText(outputString);
            break;

          case 3: // Dialog
            Rhino.UI.Dialogs.ShowTextDialog(outputString, EnglishName);
            break;
        }
      }

      return Result.Success;
    }

    /// <summary>
    /// Prompts the user for the name of a file to open
    /// </summary>
    private Result GetOpenFileName(string prompt, ref string fileName)
    {
      Result rc = Result.Cancel;

      if (string.IsNullOrEmpty(prompt))
        prompt = "File name";

      GetString gs = new GetString();
      gs.SetCommandPrompt(prompt);
      gs.AddOption(new Rhino.UI.LocalizeStringPair("Browse", "Browse"));

      if (!string.IsNullOrEmpty(fileName))
        gs.SetDefaultString(fileName);

      GetResult res = gs.Get();

      if (res == GetResult.String)
      {
        fileName = gs.StringResult();
        rc = Result.Success;
      }
      else if (res == GetResult.Option)
      {
        OpenFileDialog fileDialog = new OpenFileDialog();
        fileDialog.Filter = "RhinoScript Files (*.rvb)|*.rvb";
        fileDialog.Title = "Select RhinoScript File";
        if (fileDialog.ShowDialog() == DialogResult.OK)
        {
          fileName = fileDialog.FileName;
          rc = Result.Success;
        }
      }

      return rc;
    }

    /// <summary>
    /// Prompts the user for the name of a file to save
    /// </summary>
    private Result GetSaveFileName(string prompt, ref string fileName)
    {
      Result rc = Result.Cancel;

      if (string.IsNullOrEmpty(prompt))
        prompt = "File name";

      GetString gs = new GetString();
      gs.SetCommandPrompt(prompt);
      gs.AddOption(new Rhino.UI.LocalizeStringPair("Browse", "Browse"));

      if (!string.IsNullOrEmpty(fileName))
        gs.SetDefaultString(fileName);

      GetResult res = gs.Get();

      if (res == GetResult.String)
      {
        fileName = gs.StringResult();
        rc = Result.Success;
      }
      else if (res == GetResult.Option)
      {
        SaveFileDialog fileDialog = new SaveFileDialog();
        fileDialog.Filter = "Text Documents|*.txt";
        fileDialog.Title = "Save As";
        if (fileDialog.ShowDialog() == DialogResult.OK)
        {
          fileName = fileDialog.FileName;
          rc = Result.Success;
        }
      }

      return rc;
    }

    /// <summary>
    /// Encrypts a string with a password
    /// </summary>
    private string Encrypt(string text, string password)
    {
      if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(password))
        return null;

      string encryptedText = null;

      try
      {
        byte[] buffer = System.Text.Encoding.Unicode.GetBytes(text);
        byte[] encryptedBuffer = Encrypt(buffer, password);
        encryptedText = Convert.ToBase64String(encryptedBuffer);
      }

      catch
      {
        encryptedText = null;
      }

      return encryptedText;
    }

    /// <summary>
    /// Encrypts a byte array with a password
    /// </summary>
    private byte[] Encrypt(byte[] buffer, string password)
    {
      if (null == buffer || 0 == buffer.Length || string.IsNullOrEmpty(password))
        return null;

      byte[] encryptedBuffer = null;
      CryptoStream cryptoStream = null;

      try
      {
        Rfc2898DeriveBytes secretKey = new Rfc2898DeriveBytes(password, _keySalt);
        RijndaelManaged cipher = new RijndaelManaged();
        ICryptoTransform transform = cipher.CreateEncryptor(secretKey.GetBytes(32), secretKey.GetBytes(16));
        MemoryStream memoryStream = new MemoryStream();
        cryptoStream = new CryptoStream(memoryStream, transform, CryptoStreamMode.Write);
        cryptoStream.Write(buffer, 0, buffer.Length);
        cryptoStream.FlushFinalBlock();
        encryptedBuffer = memoryStream.ToArray();
      }

      catch
      {
        encryptedBuffer = null;
      }

      finally
      {
        if (null != cryptoStream)
          cryptoStream.Close();
      }

      return encryptedBuffer;
    }

    /// <summary>
    /// Key salt to use to derive the encryption key 
    /// </summary>
    private byte[] _keySalt = new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 };
  }
}
