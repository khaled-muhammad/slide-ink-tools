# <img src="docs/images/slide-ink-tools.svg" width="56" height="56" alt=""> Slide Ink Tools for PowerPoint

Write, highlight, and erase directly on a running PowerPoint slide. Use the floating toolbar to navigate, add blank slides, and keep annotations as editable PowerPoint shapes when the show ends.

**Windows 11 · 64-bit desktop PowerPoint · C# COM add-in prototype**

[Download installer](#install) · [Toolbar](#toolbar) · [Samsung WAF setup](#samsung-waf-setup) · [Development](#development)

![A PowerPoint slideshow with the expanded annotation toolbar, a yellow highlight, and a blue pen circle and arrow](docs/images/presentation.png)

## Install

The installer is for Windows 11 on x64 PCs with **64-bit desktop PowerPoint**. To check PowerPoint’s bitness, open **File → Account → About PowerPoint**. [Check your Office version](https://support.microsoft.com/en-us/office/lifecycle/lc-account/about-office-what-version-of-office-am-i-using)

1. Open the repository’s **Releases** tab and download `SlideInkTools-<version>-x64.msi` from the latest release’s **Assets** section.
2. Close PowerPoint and run the MSI. It installs the add-in for your Windows account without administrator privileges.
3. Open a scratch presentation and press **F5**. The floating toolbar appears in **Browse** mode. Choose **Pen** and draw with a mouse, a pen, or touch.
4. Tap **Exit presentation** or press **Esc**. Choose **Yes** to keep your marks as editable shapes, then save the deck in PowerPoint.

Install a newer MSI to update. To remove the add-in, close PowerPoint and uninstall **Slide Ink Tools for PowerPoint** under **Settings → Apps → Installed apps**.

Releases include `SHA256SUMS.txt` for verifying downloads. The MSI is currently unsigned, so Windows may display an unknown publisher notice.

## Toolbar

![The compact toolbar, the expanded toolbar with pen widths and colors, and the circular launcher](docs/images/toolbar-modes.png)

Drag the six-dot handle to move the toolbar. The chevron opens the second row. Use **×** to collapse the toolbar to its circular launcher, then tap the launcher to reopen it. The add-in remembers placement per display when Windows can identify the display.

| Tool | What it does |
| --- | --- |
| **Pen** and **Highlight** | Draw opaque ink or highlight text with translucent color. Expand the toolbar to choose a color or set the pen to **1, 3, or 6 pt**. The width applies to new pen strokes. |
| **Eraser**, **Undo**, and **Clear** | Remove whole strokes, undo changes, or clear this show’s ink from the current slide. |
| **Browse**, **Prev**, and **Next** | Return touch and mouse input to PowerPoint or navigate from the toolbar. Keyboard navigation stays available. |
| **New slide** | Insert a blank slide after the current one and jump to it. Use a standard all-slide show for this action. |
| **Blank** | Black out the slideshow temporarily. Tap it again to restore the slide. |
| **Exit presentation** | End the show with the door-and-arrow button. If you drew marks, the add-in asks whether to keep them. |

Hover over any icon with a mouse to see its name. Ink stays in memory until the show ends. Once kept and saved, the marks can be edited in PowerPoint on a computer without the add-in.

## Samsung WAF setup

1. Connect the board to the PC using **HDMI** for the display and the board’s **USB Touch Out** for touch input. Start a slideshow on the board. [Samsung WAF port specifications](https://image-us.samsung.com/SamsungUS/home/pdf/WAF_Interactive_Leaflet_V3_2_1.pdf)
2. Expand the toolbar and tap the **Palm setup** hand icon. Follow the prompts to draw one short stroke with the pen tip, one with a finger, and one with your palm. The setup samples are omitted from the slide.
3. When **Palm on** appears, swipe across an ink stroke with your palm to erase it. A stationary palm should leave strokes intact. Use **Undo** to reverse a swipe or tap the hand icon to turn palm erase off.

Palm erase depends on the contact widths Windows receives from the board. If calibration shows **Palm unavailable**, use the regular **Eraser**. Calibration is saved by display and resolution when Windows can uniquely identify the display. In mirrored mode without a unique display identity, placement and calibration apply to the current show.

## Development

- [`src/`](src/) contains the C# COM add-in, WinForms toolbar, Windows pointer input, and PowerPoint shape export.
- [`scripts/`](scripts/) contains the build and per-user registration scripts.
- [`installer/`](installer/) defines the x64 Windows Installer package.
- [Design and roadmap](PLAN.md) covers upcoming work, including whiteboard pages and ink recovery.

### Build from source

Use 64-bit Windows PowerShell with the .NET Framework compiler. The build finds the Office 15 interop assemblies in the Visual Studio **Microsoft 365 development** workload or the desktop Office interop assemblies in the Windows Global Assembly Cache. [Office interop assembly locations](https://learn.microsoft.com/en-us/visualstudio/vsto/office-primary-interop-assemblies?view=visualstudio)

Close PowerPoint and run these commands from the repository root.

```powershell
.\scripts\Build.ps1
.\scripts\Register.ps1
```

After rebuilding, start a scratch show, draw with two pen widths, and navigate away and back. Exit, keep the marks, and reopen the saved deck to confirm the shapes are editable. To remove the source registration, close PowerPoint and run `.\scripts\Register.ps1 -Unregister`. Remove the source registration before switching to an installed copy, or uninstall the MSI before registering a development build.

### Build and publish an installer

Install WiX Toolset 3.14.1 to build an MSI locally. If the tools are in a portable folder, pass its path with `-WixBin`. The MSI and checksum are written to `bin/release/`. [WiX Toolset 3.14.1](https://github.com/wixtoolset/wix3/releases/tag/wix3141rtm)

```powershell
.\scripts\Build-Installer.ps1 -Version 0.1.0
```

The [Windows installer workflow](.github/workflows/windows-installer.yml) builds and tests the MSI for pull requests and pushes to `main`. Each run provides a `windows-installer` artifact under **Actions**. To publish a release, push a version tag from the commit to release.

```powershell
git tag -a v0.1.0 -m "Slide Ink Tools 0.1.0"
git push origin v0.1.0
```

A successful tag build publishes the MSI and its SHA-256 checksum to **Releases**. Use a new `vMAJOR.MINOR.PATCH` tag for each release.

### Troubleshooting

| Symptom | Check |
| --- | --- |
| The build cannot find an Office interop DLL | Check that `office.dll` and `Microsoft.Office.Interop.PowerPoint.dll` are available in the Visual Studio Office 15 interop folder or Windows Global Assembly Cache. |
| The toolbar does not appear during a show | Reopen PowerPoint after installation. In **File → Options → Add-ins**, open **Manage → COM Add-ins → Go** and check that **Slide Ink Tools** is enabled. |
| The display works but touch does not | Check the USB Touch Out connection and verify that touch reaches Windows before starting the show. |
