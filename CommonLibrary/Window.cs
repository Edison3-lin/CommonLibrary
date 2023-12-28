/*
* Common.cs - Common API for test items
* Window.cs - Common operations for specific window
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Jacky Kao   <Jacky.Kao@quantatw.com>
*/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace CaptainWin.CommonAPI {
  /// <summary>
  /// Thie class contains operations for specific window
  /// </summary>
  public class Window {
    [DllImport("user32.dll")]
    private static extern void SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    private static extern void mouse_event(int dwFlags, int dx, int dy, 
                                           int dwData, int dwExtraInfo);
    
    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, 
                                           uint dwFlags, UIntPtr dwExtraInfo);
    
    [DllImport("user32.dll")]
    private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, 
                                              string lpszClass, string lpszWindow);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hwnd, ref Rectangle rectangle);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

    private const int _buttonDown = 0x0002;
    private const int _buttonUp = 0x0004;
    private const int _sysCommand = 0x0112;
    private const int _wnClose = 0x0010;
    private const int _wnMinimize = 0xF020;
    private const int _wnMaximize = 0xF030;

    /// <summary>
    /// Check if a specific window exists
    /// </summary>
    /// <param name="window">Identifier of the title or the ID of the window</param>
    /// <param name="isTitle">True if the identifier is the title of the window,
    /// false if the identifier is the ID of the window</param>
    /// <returns>True if the window exists, false if not</returns>
    public bool Window_Exists(string window, bool isTitle) {
      try {
        IntPtr hwnd = isTitle ? FindWindow(null, window) : new IntPtr(int.Parse(window));
        return hwnd != IntPtr.Zero;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return false;
      }
    }

    /// <summary>
    /// Focus on a specific window
    /// </summary>
    /// <param name="window">Identifier of the title or the ID of the window</param>
    /// <param name="isTitle">True if the identifier is the title of the window,
    /// false if the identifier is the ID of the window</param>
    /// <returns>0 if success, -1 if error</returns>
    public int Window_Focus(string window, bool isTitle) {
      try {
        IntPtr hwnd = isTitle ? FindWindow(null, window) : new IntPtr(int.Parse(window));
        SetForegroundWindow(hwnd);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Close a specific window
    /// </summary>
    /// <param name="window">Identifier of the title or the ID of the window</param>
    /// <param name="isTitle">True if the identifier is the title of the window,
    /// false if the identifier is the ID of the window</param>
    /// <returns>0 if success, -1 if error</returns>
    public int Window_Close(string window, bool isTitle) {
      try {
        IntPtr hwnd = isTitle ? FindWindow(null, window) : new IntPtr(int.Parse(window));
        SendMessage(hwnd, _wnClose, 0, 0);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Minimize a specific window
    /// </summary>
    /// <param name="window">Identifier of the title or the ID of the window</param>
    /// <param name="isTitle">True if the identifier is the title of the window,
    /// false if the identifier is the ID of the window</param>
    /// <returns>0 if success, -1 if error</returns>
    public int Window_Minimize(string window, bool isTitle) {
      try {
        IntPtr hwnd = isTitle ? FindWindow(null, window) : new IntPtr(int.Parse(window));
        SendMessage(hwnd, _sysCommand, _wnMinimize, 0);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Maximize a specific window
    /// </summary>
    /// <param name="window">Identifier of the title or the ID of the window</param>
    /// <param name="isTitle">True if the identifier is the title of the window,
    /// false if the identifier is the ID of the window</param>
    /// <returns>0 if success, -1 if error</returns>
    public int Window_Maximize(string window, bool isTitle) {
      try {
        IntPtr hwnd = isTitle ? FindWindow(null, window) : new IntPtr(int.Parse(window));
        SendMessage(hwnd, _sysCommand, _wnMaximize, 0);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Find a button on a specific window with its title
    /// </summary>
    /// <param name="window">Identifier of the title or the ID of the window</param>
    /// <param name="isTitleW">True if the identifier is the title of the window,
    /// false if the identifier is the ID of the window</param>
    /// <param name="button">Identifier of the title or the ID of the button</param>
    /// <param name="isTitleB">True if the identifier is the title of the button,
    /// false if the identifier is the ID of the button</param>
    /// <returns>the ID of the button if success, -1 if error</returns>
    public int Button_Find(string window, bool isTitleW, string button, bool isTitleB) {
      try {
        IntPtr hwnd = isTitleW ? FindWindow(null, window) : new IntPtr(int.Parse(window));
        IntPtr hwndButton = isTitleB ? FindWindowEx(hwnd, IntPtr.Zero, "Button", button) : new IntPtr(int.Parse(button));
        return (int)hwndButton;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }
    
    /// <summary>
    /// Click a button on a specific window with its title or ID
    /// </summary>
    /// <param name="window">Identifier of the title or the ID of the window</param>
    /// <param name="isTitleW">True if the identifier is the title of the window,
    /// false if the identifier is the ID of the window</param>
    /// <param name="button">Identifier of the title or the ID of the button</param>
    /// <param name="isTitleB">True if the identifier is the title of the button,
    /// false if the identifier is the ID of the button</param>
    /// <param name="time">Time to hold the button in seconds</param>
    /// <returns>0 if success, -1 if error</returns>
    public int Button_Click(string window, bool isTitleW, string button, bool isTitleB, int time) {
      try {
        IntPtr hwnd = isTitleW ? FindWindow(null, window) : new IntPtr(int.Parse(window));
        IntPtr hwndButton = isTitleB ? FindWindowEx(hwnd, IntPtr.Zero, "Button", button) : new IntPtr(int.Parse(button));
        SetForegroundWindow(hwnd);
        SetForegroundWindow(hwndButton);
        SendMessage(hwndButton, _buttonDown, 0, 0);
        Thread.Sleep(time * 1000);
        SendMessage(hwndButton, _buttonUp, 0, 0);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Move mouse to x,y on a specific window
    /// </summary>
    /// <param name="window">Identifier of the title or the ID of the window</param>
    /// <param name="isTitle">True if the identifier is the title of the window,
    /// false if the identifier is the ID of the window</param>
    /// <param name="x">x coordinate of the cursor</param>
    /// <param name="y">y coordinate of the cursor</param>
    /// <returns>0 if success, -1 if error</returns>
    public int Mouse_Move (string window, bool isTitle, int x, int y) {
      try {
        Thread.Sleep(50);
        IntPtr hwnd = isTitle ? FindWindow(null, window) : new IntPtr(int.Parse(window));
        SetForegroundWindow(hwnd);
        Rectangle rect = new Rectangle();
        GetWindowRect(hwnd, ref rect);
        SetCursorPos(rect.X + x, rect.Y + y);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Click mouse on a specific window
    /// </summary>
    /// <param name="window">Identifier of the title or the ID of the window</param>
    /// <param name="isTitle">True if the identifier is the title of the window,
    /// false if the identifier is the ID of the window</param>
    /// <param name="x">x coordinate of the cursor</param>
    /// <param name="y">y coordinate of the cursor</param>
    /// <param name="time">Time to hold the mouse button in seconds</param>
    /// <returns>0 if success, -1 if error</returns>
    public int Mouse_Click(string window, bool isTitle, int x, int y, int time) {
      try {
        Thread.Sleep(50);
        IntPtr hwnd = isTitle ? FindWindow(null, window) : new IntPtr(int.Parse(window));
        SetForegroundWindow(hwnd);
        Rectangle rect = new Rectangle();
        GetWindowRect(hwnd, ref rect);
        SetCursorPos(rect.X + x, rect.Y + y);
        mouse_event(_buttonDown, 0, 0, 0, 0);
        Thread.Sleep(time * 1000);
        mouse_event(_buttonUp, 0, 0, 0, 0);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }
  }
}