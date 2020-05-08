![IVANDEMES.com](https://www.ivandemes.com/wp-content/uploads/2018/09/cropped-ivandemes-header-metropolis.png)<hr>
## Convert-LNK2DEMXML
#### Description
Create VMware Dynamic Environment Manager (XML) configuration files based on existing shortcuts (LNK).

#### Getting started
Open PowerShell and run the following script.

```powershell
.\Convert-LNK2DEMXML.ps1
```

-- or --

Right-click the script in Windows Explorer and click *Run with PowerShell*.

![Right-click the script in Windows Explorer and click "Run with PowerShell".](https://www.ivandemes.com/wp-content/uploads/2020/05/explorer_convert_lnk2demxml.png)

Select the source folder that contains the shortcuts (LNK files) you want to convert.

![Select the source folder that contains the shortcuts (LNK files) you want to convert.](https://www.ivandemes.com/wp-content/uploads/2020/05/Convert-LNK2DEMXML_Source.png)

Select the target folder where you want to save the DEM XML files.

![Select the target folder where you want to save the DEM XML files.](https://www.ivandemes.com/wp-content/uploads/2020/05/Convert-LNK2DEMXML_Target.png)

Copy the XML files to your DEM configuration share (one by one, or all in once).

![Copy the XML files to your DEM configuration share (one by one, or all in once).](https://www.ivandemes.com/wp-content/uploads/2020/05/Convert-LNK2DEMXML_Result.png)

The location you want to copy the DEM XML files to is typically something like *\\\\server\demconfigshare$\general\FlexRepository\Shortcut*.

I recommend to test the shortcuts in a test or acceptance environment first, before putting in them in production.
