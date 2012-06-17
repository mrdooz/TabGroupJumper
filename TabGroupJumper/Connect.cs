/*
 * TabGroupJumper
 * Magnus Osterlind, magnus.osterlind@gmail.com
 * 
 * Provides keyboard shortcuts to be able to navigate between tab groups.
 * 
 */

using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Resources;
using System.Reflection;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;

namespace TabGroupJumper
{
  /// <summary>The object for implementing an Add-in.</summary>
  /// <seealso class='IDTExtensibility2' />
  public class Connect : IDTExtensibility2, IDTCommandTarget
  {

    static string CommandPrefix = "TabGroupJumper.Connect";
    static string[] Commands = { "JumpLeft", "JumpRight", "JumpUp", "JumpDown" };

    /// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
    public Connect()
    {
    }

    /// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
    /// <param term='application'>Root object of the host application.</param>
    /// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
    /// <param term='addInInst'>Object representing this Add-in.</param>
    /// <seealso class='IDTExtensibility2' />
    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
      _applicationObject = (DTE2)application;
      _addInInstance = (AddIn)addInInst;
      // Acting according to http://www.mztools.com/articles/2008/MZ2008004.aspx
      if (connectMode == ext_ConnectMode.ext_cm_Startup) {
        // do nothing
      } else if (connectMode == ext_ConnectMode.ext_cm_AfterStartup) {
        Array tmp = null;
        OnStartupComplete(ref tmp);
      }
    }

    /// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
    /// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
    /// <param term='custom'>Array of parameters that are host application specific.</param>
    /// <seealso class='IDTExtensibility2' />
    public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
    {
    }

    /// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
    /// <param term='custom'>Array of parameters that are host application specific.</param>
    /// <seealso class='IDTExtensibility2' />		
    public void OnAddInsUpdate(ref Array custom)
    {
    }

    /// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
    /// <param term='custom'>Array of parameters that are host application specific.</param>
    /// <seealso class='IDTExtensibility2' />
    public void OnStartupComplete(ref Array custom)
    {
      object[] contextGUIDS = new object[] { };
      Commands2 commands = (Commands2)_applicationObject.Commands;

      //Place the command on the tools menu.
      //Find the MenuBar command bar, which is the top-level command bar holding all the main menu items:
      CommandBar menuBarCommandBar = ((CommandBars)_applicationObject.CommandBars)["MenuBar"];

      //Find the Tools command bar on the MenuBar command bar:
      CommandBarControl toolsControl = menuBarCommandBar.Controls["Tools"];
      CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;
      if (toolsPopup == null)
        return;

      var addinName = "TabGroupJumper";

      // create sub menu
      CommandBar submenu = null;
      for (int i = 1; i <= toolsPopup.CommandBar.Controls.Count; ++i) {
        var cur = toolsPopup.CommandBar.Controls[i];
        if (cur.Caption == addinName) {
          submenu = ((CommandBarPopup)cur).CommandBar;
          goto MENU_FOUND;
        }
      }
      submenu = (CommandBar)commands.AddCommandBar(addinName, vsCommandBarType.vsCommandBarTypeMenu, toolsPopup.CommandBar, 1);
    MENU_FOUND:

      var jumpCommands = new[] {
        new { cmd = "JumpRight", buttonText = "Jump Right", toolTip = "Jumps to the right tab group", icon = 39, binding = "Global::Ctrl+Alt+Right Arrow" },
        new { cmd = "JumpLeft", buttonText = "Jump Left", toolTip = "Jumps to the left tab group", icon = 41, binding = "Global::Ctrl+Alt+Left Arrow" },
        new { cmd = "JumpUp", buttonText = "Jump Up", toolTip = "Jumps to tab group above", icon = 38, binding = "Global::Ctrl+Alt+Up Arrow" },
        new { cmd = "JumpDown", buttonText = "Jump Down", toolTip = "Jumps to the tab group below", icon = 40, binding = "Global::Ctrl+Alt+Down Arrow" },
      };

      int idx = 1;
      foreach (var cmd in jumpCommands) {

/*
        // remove previous version of the command (only really useful when debugging)
        try {
          commands.Item(CommandPrefix + "." + (String)cmd.cmd, 0).Delete();
        } catch (System.ArgumentException) {
        }
*/
        try {
          Command c = commands.AddNamedCommand2(_addInInstance, cmd.cmd, cmd.buttonText, cmd.toolTip, true, cmd.icon,
            ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled,
            (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

          if (c != null) {
            c.Bindings = cmd.binding;
            c.AddControl(submenu, idx++);
          }

        } catch (System.ArgumentException) {
          //If we are here, then the exception is probably because a command with that name
          //  already exists. If so there is no need to recreate the command and we can 
          //  safely ignore the exception.
        }
      }
    }

    /// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
    /// <param term='custom'>Array of parameters that are host application specific.</param>
    /// <seealso class='IDTExtensibility2' />
    public void OnBeginShutdown(ref Array custom)
    {
    }

    /// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
    /// <param term='commandName'>The name of the command to determine state for.</param>
    /// <param term='neededText'>Text that is needed for the command.</param>
    /// <param term='status'>The state of the command in the user interface.</param>
    /// <param term='commandText'>Text requested by the neededText parameter.</param>
    /// <seealso class='Exec' />
    public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
    {
      if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone) {
        foreach (var c in Commands) {
          if (commandName == CommandPrefix + "." + c) {
            status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
            return;
          }
        }
      }
    }

    /// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
    /// <param term='commandName'>The name of the command to execute.</param>
    /// <param term='executeOption'>Describes how the command should be run.</param>
    /// <param term='varIn'>Parameters passed from the caller to the command handler.</param>
    /// <param term='varOut'>Parameters passed from the command handler to the caller.</param>
    /// <param term='handled'>Informs the caller if the command was handled or not.</param>
    /// <seealso class='Exec' />
    public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
    {
      handled = false;

      // Check if we can handle this command
      if (executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault) {
        foreach (var c in Commands) {
          if (commandName == CommandPrefix + "." + c) {
            handled = true;
            break;
          }
        }

        if (!handled)
          return;

        List<EnvDTE.Window> topLevel = new List<EnvDTE.Window>();
        bool horizontal = commandName.EndsWith("Left") || commandName.EndsWith("Right");

        // Documents with a "left" or "top" value > 0 are the focused ones in each group, 
        // so we only need to collect those
        foreach (EnvDTE.Window w in _applicationObject.Windows) {
          if (w.Kind == "Document" && (horizontal && w.Left > 0 || w.Top > 0))
            topLevel.Add(w);
        }
        if (horizontal)
          topLevel.Sort((a, b) => a.Left < b.Left ? -1 : 1);
        else
          topLevel.Sort((a, b) => a.Top < b.Top ? -1 : 1);

        // find the index of the active document
        var activeDoc = _applicationObject.ActiveDocument;
        int activeIdx = 0;
        for (int i = 0; i < topLevel.Count; ++i) {
          if (topLevel[i].Document == activeDoc) {
            activeIdx = i;
            break;
          }
        }

        // set the new active document
        if (horizontal) {
          bool jumpLeft = commandName.EndsWith("JumpLeft");
          activeIdx += jumpLeft ? -1 : 1;
        } else {
          bool jumpUp = commandName.EndsWith("JumpUp");
          activeIdx += jumpUp ? -1 : 1;
        }

        activeIdx = (activeIdx < 0 ? activeIdx + topLevel.Count : activeIdx) % topLevel.Count;
        topLevel[activeIdx].Activate();
      }
    }

    private DTE2 _applicationObject;
    private AddIn _addInInstance;
  }
}
