# excel-timecode

Microsoft Excel custom functions (JavaScript) for working with video timecode standards and wall
time durations.

**IMPORTANT**: These custom functions are only supported in the latest subscription versions of
Office for Windows, Mac, or the web (NOT in one-time purchase versions of Office 2019 or earlier).

Designed for film & television composers, though these may be useful for anyone who works with
timecode values in Google Sheets.

Primary Author: [Eric Barndollar](https://barndollarmusic.com)

This is open source software that is free to use and share, as covered by the
[MIT License](LICENSE).

# Using custom functions
The last 2 arguments to every function below are `frameRate` and `dropType` values.

Data validation list of supported `frameRate` values (see templates above for example usage):
```
23.976,24.000,25.000,29.970,30.000,47.952,48.000,50.000,59.940,60.000
```
**IMPORTANT**: The `frameRate` value must be **Plain text** type (not a number) and include exactly
2 or 3 decimal digits after a period. This is to avoid any possible confusion over *e.g.* whether
`24` means `23.976` or `24.000`.

Data validation list of `dropType` values (see templates above for example usage):
```
non-drop,drop
```

## Most common functions
```JavaScript
=TIMECODE.TC_TO_WALL_SECS("00:00:01:02", "50.00", "non-drop")
```
- Yields `1.04` secs (true seconds of wall time measured from `00:00:00:00`).

```JavaScript
=TIMECODE.WALL_SECS_BETWEEN_TCS("00:00:01:03", "00:02:05:11", "24.00", "non-drop")
```
- Yields `124.33333333...` secs (true seconds of wall time between the timecodes).

```JavaScript
=TIMECODE.WALL_SECS_TO_DURSTR(3765)
```
- Yields `"1h 02m 45s"` (a human-readable duration string). Rounds to nearest second.

```JavaScript
=TIMECODE.WALL_SECS_TO_TC_LEFT(1.041, "50.00", "non-drop")
```
- Yields `"00:00:01:02"`, the timecode of the closest frame that is exactly at or
before (*i.e.* to the left of) the given `wallSecs` value of `1.041` (true seconds of
wall time measured from `00:00:00:00`).

```JavaScript
=TIMECODE.WALL_SECS_TO_TC_RIGHT(1.041, "50.00", "non-drop")
```
- Yields `"00:00:01:03"`, the timecode of the closest frame that is exactly at or
after (*i.e.* to the right of) the given `wallSecs` value of `1.041` (true seconds of
wall time measured from `00:00:00:00`).

## Other functions (more advanced)
```JavaScript
=TIMECODE.TC_ERROR("01:02:03:04", "23.976", "non-drop")
```
- Yields an error string if timecode (or format) is invalid, or an empty string otherwise.

```JavaScript
=TIMECODE.TC_TO_FRAMEIDX("00:00:01:02", "50.00", "non-drop")
```
- Yields `52` (the timecode refers to the 53rd frame of video, counting from `00:00:00:00` as
index 0). Dropped frames are not given index values (so in 29.97 drop, `00:00:59:29` has index
`1799` and `00:01:00:02` has index `1800`).

```JavaScript
=TIMECODE.FRAMEIDX_TO_TC(52, "50.00", "non-drop")
```
- Yields `"00:00:01:02"`, the timecode of the given frame index.

```JavaScript
=TIMECODE.FRAMEIDX_TO_WALL_SECS(52, "50.00", "non-drop")
```
- Yields `1.04` secs (true seconds of wall time measured from `00:00:00:00`).

```JavaScript
=TIMECODE.WALL_SECS_TO_FRAMEIDX_LEFT(1.041, "50.00", "non-drop")
```
- Yields `52`, the frame index of the closest frame that is exactly at or
before (*i.e.* to the left of) the given `wallSecs` value of `1.041` (true seconds of
wall time measured from `00:00:00:00`).

```JavaScript
=TIMECODE.WALL_SECS_TO_FRAMEIDX_RIGHT(1.041, "50.00", "non-drop")
```
- Yields `53`, the frame index of the closest frame that is exactly at or
after (*i.e.* to the right of) the given `wallSecs` value of `1.041` (true seconds of
wall time measured from `00:00:00:00`).

# Contributing Code

For the custom functions themselves, please first update and test your changes to this
repository: [gsheets-timecode](https://github.com/barndollarmusic/gsheets-timecode).

If you make changes to [manifest.xml](manifest.xml), see
[Clear the Office Cache](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache).

# Testing with Excel Locally

Open this project in Visual Studio Code.
- **Terminal > Run Task... > Watch**
- **Terminal > Run Task... > Dev Server**
- **View > Run**, then select **Excel Desktop (Edge Chromium)** and hit play button (`F5`).

Or from the command line (run these commands in separate tabs):
```bash
npm run watch
```

```bash
npm run dev-server
```

```bash
npm run start:desktop
```