# Interactive Display PowerPoint Add-in Plan

This document records the design and roadmap. See the [README](README.md) for the features available in the current prototype and how to run it.

Build a **Windows desktop PowerPoint add-in with a floating toolbar on the Samsung WAF**. Connect HDMI for video and USB Touch Out for input, start PowerPoint normally with F5 or Shift+F5, and use the toolbar directly on the board. At the end of the show, save annotations as editable objects on the original slides or in an annotated copy of the deck.

The first milestone is a hardware prototype. Samsung specifies a passive pen and recognition of different contact sizes, including an eraser. The signals that reach Windows over USB determine whether palm erase and automatic pen-tip switching will work on the exact board. [Samsung WAF specifications](https://image-us.samsung.com/SamsungUS/home/pdf/WAF_Interactive_Leaflet_V3_2_1.pdf)

## Presenting experience

| Feature | Planned behavior |
|---|---|
| **Floating toolbar** | Use the compact icon strip and expand it for additional tools. Drag the handle to reposition it. Close it to a translucent circular launcher that can also be dragged and tapped to reopen. It remembers placement per display. The layout can later support edge snapping, vertical orientation, and a shortcut to hide or reopen the whole surface. |
| **Pen and highlighter** | Pick colors and thickness from the toolbar. The pen draws opaque ink, and the highlighter stays translucent over slide content. The WAF pen works in the selected drawing tool even if Windows reports it as ordinary touch. |
| **Eraser and palm swipe** | The eraser removes the portion of a mark you sweep across. On a validated WAF setup, a deliberate palm swipe temporarily activates that eraser and lifting your hand restores the previous tool. One swipe is one Undo action. The toolbar eraser is always available. |
| **Undo, Redo, and Clear** | Undo and redo the current page’s changes. Clear removes all marks made on that slide or whiteboard during the current session and can be undone in one step. |
| **Navigate** | Switch back to normal PowerPoint interaction for links, videos, and touch navigation. Add Previous and Next buttons on the board, and keep the keyboard and presenter remote usable while drawing. |
| **New slide** | Insert a blank slide after the current slide from the toolbar in a standard slide show. PowerPoint jumps to the new slide and the add-in tracks annotations by slide ID. |
| **Blank screen and whiteboard** | **Blackout** temporarily hides the slide and restores it with one tap. **Whiteboard** opens a blank writing surface and returns you to the same slide when you close it. You can keep a whiteboard page as a new slide. |
| **End & Save** | Save to the current deck, save an annotated copy, or discard the session. Ending with Esc also brings up the save choices after PowerPoint returns to its editing view. |

## Technical design

### PowerPoint integration

Use a C# VSTO add-in targeting .NET Framework 4.8, with a ribbon command for settings and recovery. The add-in listens for slideshow start, slide changes, and show end. The development machine has 64-bit desktop PowerPoint. Target that installation, Windows 11, local `.pptx` files, and the exact WAF model for the pilot. Use Windows **Extend** mode with Presenter View on the laptop. Test Duplicate mode separately. [Microsoft’s PowerPoint VSTO walkthrough](https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-powerpoint?view=visualstudio) · [VSTO runtime support](https://learn.microsoft.com/en-us/visualstudio/vsto/visual-studio-tools-for-office-runtime?view=visualstudio) · [PowerPoint Presenter View](https://support.microsoft.com/en-us/powerpoint/present-your-slide-show)

### Windows overlay and input

Use WPF for the toolbar and a native Windows surface for drawing and pointer input. The toolbar follows the audience slideshow window. The ink surface tracks the actual slide rectangle, including letterboxing, and receives touch only in a drawing mode. In Navigate mode, taps reach PowerPoint. Bind each stroke to the current slide when it starts. Pause capture when PowerPoint changes slides or the monitor mapping becomes uncertain, then resume once the new target is confirmed. This protects against ink landing on the wrong slide during transitions or cable reconnects. [PowerPoint slide-change event](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.slideshownextslide) · [Windows pointer input](https://learn.microsoft.com/en-us/windows/win32/inputmsg/functions) · [Windows per-monitor DPI](https://learn.microsoft.com/en-us/windows/win32/hidpi/high-dpi-desktop-application-development-on-windows)

### Palm recognition

Build a diagnostic that records how Windows classifies the front tip, back tip, fingertip, resting hand, and intentional palm swipe. If Windows exposes a reliable eraser signal, use it. If it exposes meaningful touch size or a consistent contact pattern, calibrate a deliberate moving-palm gesture and show an eraser preview before removing ink. A resting hand or ambiguous contact should leave the marks intact. If that WAF firmware exposes indistinguishable contacts, ship pen and manual erasing on that setup and leave automatic palm erase disabled. [Windows touch contact data](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-pointer_touch_info) · [Windows pen eraser flags](https://learn.microsoft.com/en-us/windows/win32/inputmsg/pen-flags-constants)

### Saving and recovery

During the show, store completed drawing actions in a recoverable local draft and render from that draft. After the show, write the final marks as tagged PowerPoint vector shapes on the slides they belong to, using stable slide IDs so reordering the deck doesn’t misplace ink. Someone opening the saved deck can see and edit the marks in PowerPoint without installing the add-in. For v1, palm erase and Undo apply to the current show’s ink. Marks saved in earlier sessions can be edited as shapes in PowerPoint’s Normal view. If a save fails or a slide was deleted, retain the affected draft and report what still needs saving. Retrying should produce one copy of each annotation. [PowerPoint slide IDs](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.slideid) · [Editable freeform shapes](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.buildfreeform) · [PowerPoint SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.savecopyas)

## Build order and release checks

| Milestone | What must work before moving on |
|---|---|
| **1. Hardware and PowerPoint prototype** | On the WAF, draw a line over a live slideshow using the supplied pen. Confirm it lands where the tip touches, keyboard navigation still works, the toolbar stays on the audience display, and the line can be saved and reopened as an editable shape. Record which palm and pen signals Windows actually delivers. |
| **2. Presenting tools** | Build the movable toolbar, pen, highlighter, erase, Undo/Redo, and navigation. Test a 4:3 deck on the 16:9 board, mixed display scaling, slide transitions, and going back to earlier slides. Measure pen-to-ink delay on the board at 4K and 60 Hz. |
| **3. Whiteboard and saving** | Add blackout, whiteboard pages, draft recovery, and the save choices. Reopen saved decks on a machine without the add-in. Test an interrupted save, a read-only deck, and a restart with unsaved annotations. |
| **4. Room pilot and installer** | Run a full presentation with the laptop, WAF, USB cable, and presenter remote. Test reconnecting HDMI and USB mid-show, switching apps, and closing and reopening the toolbar. For palm erase, require reliable deliberate swipes and zero unwanted erasures in a scripted set of ordinary writing and resting-hand trials. Package a signed installer and verify install and uninstall on a clean Windows machine. [Microsoft’s VSTO installer guidance](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution-by-using-windows-installer) |

## Estimate and current status

Allow **5–7 weeks for one developer** to reach a pilot on one WAF setup, assuming access to the physical board in week one. Spend the first **3–5 working days** proving the whole workflow by drawing on a live slide with the WAF pen, navigating away and back, saving, and reopening the deck.

The prototype uses a C# COM add-in and WinForms toolbar because the development machine has the PowerPoint interop libraries but lacks the VSTO build workload. It exercises the same PowerPoint slideshow events and writes editable shapes. The first room test confirmed pen, finger, and palm interaction on the Samsung WAF after the rendering update. The add-in now remembers the toolbar position per display and successful palm calibration per display and screen resolution. See [README.md](README.md) for the current build and trial instructions.
