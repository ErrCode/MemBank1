# Convert Macrofun.hlp to Macrofun.chm

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

I went with option 1, but I could have gone with option 2. Both would require doing hacks/using questionable software components. I did reason that option 2 meant using an executable that Microsoft specifically stopped supporting \(probably because of the potential exploits that `.HLP` files permitted\) for a lot longer duration than option 1, but in hindsight 



