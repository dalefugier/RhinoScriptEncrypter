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
    /// <summary>
    /// The one and only instance of this command
    /// </summary>
    static RhinoScriptEncrypterCommand _instance;

    /// <summary>
    /// Public constructor
    /// </summary>
    public RhinoScriptEncrypterCommand()
    {
      _instance = this;
    }

    /// <summary>
    /// Returns the one and only instance of this command
    /// </summary>
    public static RhinoScriptEncrypterCommand Instance
    {
      get { return _instance; }
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
      string fileName = string.Empty;

      // Prompt or a filename
      if (mode == RunMode.Interactive)
      {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Filter = "RhinoScript Files (*.rvb)|*.rvb";
        openFileDialog.Title = "Select RhinoScript File";
        if (openFileDialog.ShowDialog() != DialogResult.OK)
          return Result.Cancel;

        fileName = openFileDialog.FileName;
      }
      else
      {
        Result rc = Rhino.Input.RhinoGet.GetString("RhinoScript file to encrypt", false, ref fileName);
        if (rc != Rhino.Commands.Result.Success)
          return rc;
      }

      // Verify the filename
      fileName = fileName.Trim();
      if (string.IsNullOrEmpty(fileName))
        return Result.Nothing;

      // Verify the file
      if (!File.Exists(fileName))
      {
        string error = string.Format("RhinoScript file not found - {0}\n", fileName);
        if (mode == RunMode.Interactive)
          MessageBox.Show(error, EnglishName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        else
          RhinoApp.WriteLine(error);
        return Result.Failure;
      }

      // Generate an encryption password
      string encryptPassword = Guid.NewGuid().ToString();

      // Do the encryption
      string encryptedFileName = string.Empty;
      try
      {
        // Read the script
        string clearString = File.ReadAllText(fileName);
        // Encrypt the script
        string encryptedString = Encrypt(clearString, encryptPassword);
        // Write the encrypted script
        encryptedFileName = Path.ChangeExtension(fileName, ".rvbx");
        File.WriteAllText(encryptedFileName, encryptedString);
      }
      catch (Exception ex)
      {
        string error = string.Format("Error encrypting RhinoScript file - {0}\n{1}\n", fileName, ex.Message);
        if (mode == RunMode.Interactive)
          MessageBox.Show(error, EnglishName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        else
          RhinoApp.WriteLine(error);
        return Result.Failure;
      }

      // Report the results
      string output = "RhinoScript file encryption successful!\n";
      output += string.Format("  Input filename: {0}\n", fileName);
      output += string.Format("  Output filename: {0}\n", encryptedFileName);
      output += string.Format("  Decrypt password: {0}\n", encryptPassword);

      if (mode == RunMode.Interactive)
        Rhino.UI.Dialogs.ShowTextDialog(output, EnglishName);
      else
        RhinoApp.WriteLine(output);

      return Result.Success;
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
