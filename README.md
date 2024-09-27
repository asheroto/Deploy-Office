[![Release](https://img.shields.io/github/v/release/asheroto/Deploy-Office)](https://github.com/asheroto/Deploy-Office/releases)
[![GitHub Release Date - Published_At](https://img.shields.io/github/release-date/asheroto/Deploy-Office)](https://github.com/asheroto/Deploy-Office/releases)
[![GitHub Sponsor](https://img.shields.io/github/sponsors/asheroto?label=Sponsor&logo=GitHub)](https://github.com/sponsors/asheroto?frequency=one-time&sponsor=asheroto)
<a href="https://ko-fi.com/asheroto"><img src="https://ko-fi.com/img/githubbutton_sm.svg" alt="Ko-Fi Button" height="20px"></a>
<a href="https://www.buymeacoffee.com/asheroto"><img src="https://img.buymeacoffee.com/button-api/?text=Buy me a coffee&emoji=&slug=Deploy-Office&button_colour=FFDD00&font_colour=000000&font_family=Lato&outline_colour=000000&coffee_colour=ffffff)" height="40px"></a>

# Deploy Office

![screenshot](https://github.com/asheroto/Deploy-Office/assets/49938263/d6ef4e34-7f77-46da-80cd-6b494d321fac)

> [!NOTE]
> Version 1.0 has been released!

Easily install Office 2019, 2021, 2024, or Microsoft 365. Simply open the program and it will start installing in 30 seconds, or change the options and click `Start` to run immediately.

Everything is downloaded from the cloud, so you'll always have the latest version available when installing.

Desktop shortcuts will also be created by default for any installed Office applications.

## Why Use This?

Because it's a tiny 260 KB executable that downloads the latest version of Office and installs it for you. No need to download the installer, write XML configuration files, or run the setup manually.

## Does This Legitimately Install Microsoft Office?

**Yes!** This is **NOT** a bootleg or modified version of Office.

As outlined in [How it Works](#how-it-works), the software fetches `setup.exe` directly from Microsoft.

Office still needs a valid license, which you can activate post-installation. Typically, after installing Office, you're given a trial period before you need to activate it.

This project and its creators are not affiliated with Microsoft. This software uses the official Office installer provided by Microsoft, which is publicly accessible.

Being open-source, the source code for this software accessible for anyone to read and inspect.

## Legal

All rights of Microsoft Office or any of Office products listed belong to [Microsoft Corporation](https://microsoft.com).

## Supported Editions/Products

> [!NOTE]
> If you specify a calendar year like 2021, it will limit the selection to non-365 editions only.
> Similarly, specifying the calendar year `Microsoft 365` will limit the selection to only Microsoft 365 editions.
> This is a design choice to make choosing the edition simple.

### Years

1. 2019
2. 2021
3. 2024
4. Microsoft 365

### Editions/Products

1. Microsoft 365 Family/Personal
2. Microsoft 365 Small Business
3. Microsoft 365 Education
4. Microsoft 365 Enterprise
5. Home & Business
6. Home & Student
7. Personal
8. Professional
9. Professional Plus
10. Professional Plus - Volume
11. Standard
12. Standard - Volume
13. Visio Standard
14. Visio Standard - Volume
15. Visio Professional
16. Visio Professional - Volume
17. Project Standard
18. Project Standard - Volume
19. Project Professional
20. Project Professional - Volume
21. Access
22. Access - Volume
23. Excel
24. Excel - Volume
25. Outlook
26. Outlook - Volume
27. PowerPoint
28. PowerPoint - Volume
29. Publisher
30. Publisher - Volume
31. Word
32. Word - Volume

## Specify Default Office Edition/Product Installation at Runtime (Optional)

> [!IMPORTANT]
> Starting from version 1.0, the ordering of the editions now begins at 1 instead of 0.
> `Deploy-Office.txt` is no longer supported.

This an **OPTIONAL** feature and is an optional feature to help automate the installation of Office.

When `Deploy-Office.exe` runs, it checks if the `Deploy-Office.ini` file exists, and if so, uses its settings to determine the default Office edition/product for installation.

### Setting up `Deploy-Office.ini`

1. **Open your favorite text editor** such as Notepad, Notepad++, or Visual Studio Code.
2. **Use the information & examples below to specify the default year/edition.** Additional options like creating shortcuts, excluding Teams, or excluding OneDrive are optional.
3. **Save the file** as `Deploy-Office.ini` in the same folder as `Deploy-Office.exe`.

### Examples

<details>
<summary>Simple example using index numbers (2021 Home & Business)</summary>

```ini
year=2
edition=5
```

- Specifies the year as **2021** (index 2).
- Specifies the edition to **Home & Business** (index 5).
- Creates **desktop shortcuts** (enabled by default).

</details>

<details>
<summary>Simple example using values (2024 Word)</summary>

```ini
year=2024
edition=Word
```

- Specifies the year as **2024**.
- Sets the edition to **Word**.
- Creates **desktop shortcuts** (enabled by default).

</details>

<details>
<summary>Detailed example using values and optional settings (2024 Microsoft 365 Family/Personal)</summary>

```ini
year=2024
edition=Microsoft 365 Family/Personal
shortcuts=false
exclude_teams=true
exclude_onedrive=true
```

- Specifies the year as **2024**.
- Sets the edition to **Microsoft 365 Family/Personal**.
- Disables the creation of **desktop shortcuts**.
- Excludes the installation of **Microsoft Teams**.
- Excludes the installation of **OneDrive**.

</details>

<details>
<summary>Detailed example using values and optional settings (Professional Plus - Volume)</summary>

```ini
year=2024
edition=10
shortcuts=false
exclude_teams=true
exclude_onedrive=true
```

- Specifies the year as **2024**.
- Sets the edition to **Professional Plus - Volume**.
- Disables the creation of **desktop shortcuts**.
- Excludes the installation of **Microsoft Teams**.
- Excludes the installation of **OneDrive**.

</details>

### Settings

| Name               | Required | Default                    | Description                                                                                                                 |
| ------------------ | -------- | -------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `year`             | Yes      | 2021                       | Specifies the year for the Office edition/product installation. Can be an index number or a specific year.                  |
| `edition`          | Yes      | Professional Plus - Volume | Specifies the edition of the Office product to be installed. Can be an index number or a specific edition name.             |
| `shortcuts`        | No       | Enabled                    | Determines whether desktop shortcuts should be created. Set to `true` or `false`.                                           |
| `exclude_teams`    | No       | Disabled                   | Specifies whether to exclude the installation of Microsoft Teams. Set to `true` to exclude Teams, otherwise set to `false`. |
| `exclude_onedrive` | No       | Disabled                   | Specifies whether to exclude the installation of OneDrive. Set to `true` to exclude OneDrive, otherwise set to `false`.     |

Note that shortcuts will be created by default unless disabled.

## How it Works

1. If using `Deploy-Office.ini`, it will set the default edition/product to install
2. Downloads `setup.exe` from https://officecdn.microsoft.com/pr/wsus/setup.exe
3. Extracts configuration.xml
4. Runs `setup.exe /Configure configuration.xml`

**Note**: if you need more advanced options with your Office setup, visit [config.office.com](https://config.office.com/deploymentsettings).

## Prerequisites

- Windows 10 or later
- Server 2019 or later
- .NET Framework 4.8 (which you probably already have)
- No additional dependencies needed 😊

## Download

Download the [latest version](https://github.com/asheroto/Deploy-Office/releases/latest/download/Deploy-Office.zip) in the releases section.