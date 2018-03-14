# Convert Macrofun.hlp to Macrofun.chm

ErrCode log 2018-03-01

## Problem Description

Have you downloaded the Office Excel XLL SDK \(e.g. [Excel 2010 XLL SDK](https://www.microsoft.com/en-au/download/details.aspx?id=20199)\) only to find that the help doesn't seem to give you any specific information about the parameters to supply for when trying to use`xlfInput`?

Well it turns out that `xlfInput` are amongst the legacy MACRO SHEET functions, and thus requires you

* to look up the mapping of `xlfInput` to `INPUT` \(MACRO SHEET function\) via `INTLMAP.XLS` included with the SDK; and
* to have the macro sheet function documentation for `INPUT` via `MACROFUN.HLP`

Unfortunately, the `MACROFUN.HLP` is no longer supported in Window 10, so what to do now?

## Approach

After researching into how to read a `.HLP` file on a modern Windows machine, I realised that there were two ways:

1. Convert from this old format to the newer `.CHM` format; or
2. Install the old help file viewer executable that still supported `.HLP` files

I went with option 1, but I could have gone with option 2. Both would require doing hacks/using questionable software components.

I did reason that option 2 meant two things:

* using an executable that Microsoft specifically stopped supporting \(probably because of the potential exploits that `.HLP` files permitted\) for a lot longer duration than option 1 \(considering the time I need to read the `.HLP` file\). That said, in hindsight the software used for option 1 could have had their own problems.
* as I didn't get the the `MACROFUN.HLP` file from Microsoft directly, so I couldn't know whether the file would contain anything bad that could exploit the deprecated help executable above.

Also option 1 did have the benefit of Ulrich Kulle very kindly supplying every step required on his website \[see: [Converting WinHelp \(HLP\) to HTMLHelp \(CHM\) - Table of Contents](http://www.help-info.de/en/Help_Info_WinHelp/hw_converting.htm)\].

As with everything sourced from internet, I'd trust it more if there was source code available \(shows they have less or nothing to hide\) and if installation didn't require elevated privileges.

By using 7zip to unzip almost everything \(works well on self-extracting exes and even most setup exes\), there was mostly no need to run any unnecessary setup/self-extracting exes that may have undesirable payloads included. So fortunately, most of the tools would run stand-alone without any installation and no need for elevated admin privileges.

There was one exception - [Microsoft's own HTMLHelp tool](http://msdn2.microsoft.com/en-us/library/ms669985.aspx), which if I didn't install via the setup didn't register the RTF to HTML converter required for it to work correctly. Note: when you do install HTMLHelp via the setup, it will give a complete with warning as described by U.Kulle [here](http://www.help-info.de/en/Help_Info_HTMLHelp/hh_download_install.htm). This warning can be be ignored since Windows 10 has newer help viewing components.

\***PRO TIP**\* always use [VirusTotal](https://www.virustotal.com) to check files you get from the internet before you even open them with anything.

### Going the extra mile

Following the initial steps provided by U.Kulle, I found that the output lacked images. This was partly solved by converting the `.BMP` images from the decompiling process to `.GIF` images \(and placed under the `images` folder\) as suggested on his website [here](http://www.help-info.de/en/Help_Info_WinHelp/hw_converting.htm#Updating).

Unfortunately, this only helped for the `.BMP` image types - there were more `.MRB` files in the decompile directory that I didn't know what to do with and it took quite a bit of googling to find a tool that would help convert it. The `deark` tool was capable of converting the `.MRB` files \([apparently](http://fileformats.archiveteam.org/wiki/Segmented_Hypergraphics) `.MRB` stood for multiple resolution bitmap and is related to segmented hypergraphics that had the extension`.SHG`\) and spat out three differently sized `.BMP` files for each `.MRB` file. Choosing to only convert the largest to `.GIF` and moving them to the `images` folder finally allowed all images to appear in the output `CHM`.

