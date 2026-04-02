# 🌈 
 
# OutlookEmailToMarkdown

An Outlook VBA macro that converts selected emails into clean, structured Markdown and copies them to your clipboard — ready to paste into ChatGPT, Claude, or any other AI/LLM interface. No FM20.DLL required.

## Features

- **No FM20.DLL dependency** — uses Win32 API directly for clipboard access, so no extra references or ActiveX components are needed
- **Reply chain parsing** — detects quoted reply and forward boundaries and structures them as hierarchical Markdown headings (`##`, `###`)
- **Whitespace normalization** — collapses redundant blank lines and trims trailing spaces to reduce token usage
- **Signature stripping** — automatically removes common email signature blocks ("Sent from my iPhone", "Get Outlook for", `-- ` delimiters, etc.)
- **Multi-email support** — select one or more emails and they are all converted and concatenated in a single clipboard copy
- **32-bit and 64-bit compatible** — conditional `#If VBA7` compilation handles both Outlook versions

## Output Format

Each email is structured like this:

```markdown
# Email Subject
**From:** Sender Name | **Sent:** 2026-04-01 09:30
**To:** recipient@example.com
**CC:** other@example.com

## Body

The main message content goes here...

### Quoted Reply 1

The first quoted reply in the chain...

***
```

## Installation

1. Open Outlook
2. Press **Alt+F11** to open the VBA editor
3. Go to **Insert → Module**
4. Paste the contents of `CopyOutlookDesktopMsgs.bas` into the module
5. Close the VBA editor

No additional references or libraries are required.

## Usage

1. Select one or more emails in the Outlook Explorer (mail list view)
2. Press **Alt+F11** to open the VBA editor
3. Place your cursor inside the `CopySelectedEmailsToMarkdown` subroutine
4. Press **F5** to run, or assign the macro to a Quick Access Toolbar button for one-click use
5. Paste the result (**Ctrl+V**) into your AI interface of choice

### Assigning to a Toolbar Button (Recommended)

1. Right-click the Quick Access Toolbar → **Customize Quick Access Toolbar**
2. Under "Choose commands from," select **Macros**
3. Find `CopySelectedEmailsToMarkdown` and click **Add**
4. Click **OK** — the macro now runs with one click

## Compatibility

| Environment | Supported |
|---|---|
| Outlook 2016 / 2019 / 2021 | ✅ |
| Microsoft 365 (Desktop) | ✅ |
| 64-bit Office | ✅ |
| 32-bit Office | ✅ |
| Outlook Web / OWA | ❌ |

