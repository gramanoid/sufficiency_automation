# How to Link PowerPoint Tables to Excel Data

## Overview
This guide shows how to create PowerPoint tables that automatically update when you change the Excel file. No coding required - uses built-in Microsoft Office features.

---

## Method 1: Paste Link (Recommended for Tables)

### Step-by-Step Instructions

#### In Excel:
1. Open `Haleon - 2026 MEA Budget Sufficiency_UPDATED.xlsx`
2. Go to the **2026 Sufficiency** sheet
3. Select the data range you want to link (e.g., one market's rows)
4. Press `Ctrl+C` (or `Cmd+C` on Mac) to copy

#### In PowerPoint:
1. Open your presentation
2. Navigate to the slide where you want the linked table
3. Go to **Home** > **Paste** dropdown arrow > **Paste Special**
4. Select **Paste Link**
5. Choose **Microsoft Excel Worksheet Object**
6. Click **OK**

### How Updates Work
- **When you open the PPT:** It will prompt to update links - click **Update Links**
- **Manual update:** Right-click the table > **Update Link**
- **Automatic:** Open both files, changes sync automatically

---

## Method 2: Insert Linked Object (Alternative)

### Steps:
1. In PowerPoint, go to **Insert** > **Object**
2. Select **Create from file**
3. Click **Browse** and select the Excel file
4. **Check the "Link" checkbox** (important!)
5. Click **OK**

This embeds a live Excel view that updates when the source changes.

---

## Important Tips

### File Locations
- Keep Excel and PowerPoint files in the **same folder** or a **shared network drive**
- If you move the Excel file, the link will break
- Relative paths work best when files are together

### For Sharing with Colleagues
1. Keep both files in a shared folder (OneDrive, SharePoint, network drive)
2. Colleagues should open PowerPoint from the same location
3. When prompted "Update Links?", click **Yes**

### Troubleshooting Broken Links
If links stop working:
1. In PowerPoint: **File** > **Info** > **Edit Links to Files**
2. Select the broken link
3. Click **Change Source**
4. Navigate to the new Excel file location
5. Click **Update Now**

---

## Best Practices

### For Each Market Slide:
1. Create separate linked ranges for each market in Excel
2. Copy one market's data at a time
3. Paste Link into the corresponding PPT slide
4. This keeps slides independent and easier to manage

### Naming Ranges (Optional but Helpful)
In Excel, you can name ranges for easier reference:
1. Select the data range (e.g., KSA rows)
2. Click in the **Name Box** (left of formula bar)
3. Type a name like `KSA_Data`
4. Press Enter

Named ranges make it easier to identify what's linked.

---

## Quick Reference Card

| Action | How |
|--------|-----|
| Copy Excel data | `Ctrl+C` / `Cmd+C` |
| Paste as Link | **Home > Paste > Paste Special > Paste Link** |
| Update links | Right-click table > **Update Link** |
| Fix broken links | **File > Info > Edit Links to Files** |
| Check link status | **File > Info** (shows linked files) |

---

## Files in This Folder

| File | Purpose |
|------|---------|
| `Haleon - 2026 MEA Budget Sufficiency_UPDATED.xlsx` | Source Excel (use this for linking) |
| `BACKUP_*.xlsx` | Original backup (don't modify) |
| `2026 Sufficiency_Market Deep Dive_*.pptx` | PowerPoint to update |

---

## Workflow Summary

```
1. Update Excel data
          ↓
2. Save Excel file
          ↓
3. Open PowerPoint
          ↓
4. Click "Update Links" when prompted
          ↓
5. Tables now show latest Excel data
```

No coding, no macros - just standard Office features!
