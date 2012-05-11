using System;

namespace RhinoScriptEncrypter
{
  /// <summary>
  /// RhinoScriptEncrypterPlugIn
  /// </summary>
  public class RhinoScriptEncrypterPlugIn : Rhino.PlugIns.PlugIn
  {
    /// <summary>
    /// The one and only instance of this plug-in
    /// </summary>
    static RhinoScriptEncrypterPlugIn _instance;

    /// <summary>
    /// Public constructor
    /// </summary>
    public RhinoScriptEncrypterPlugIn()
    {
      _instance = this;
    }

    /// <summary>
    /// Returns the one and only instance of this plug-in
    /// </summary>
    public static RhinoScriptEncrypterPlugIn Instance
    {
      get { return _instance; }
    }
  }
}