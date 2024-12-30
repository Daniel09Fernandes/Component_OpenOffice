# DinosOffice: Delphi Components for LibreOffice

### Important
----

the DocVisible property was included, by default it is false, so when generating the spreadsheet it is not displayed on the screen,
If you want to see the spreadsheet being generated, set it to true;

## Running in: 
### Vcl, Fmx(Win32/64), Unigui, Intraweb

![image](https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/ddec6e5c-ea7e-4840-a080-facc7e2384bf)
![image](https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/ec6738a6-f775-4d99-996e-5a6fdc092a73) 
![image](https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/d1509f79-eb7e-496a-9292-a2bfab710b7b) &nbsp;
![image](https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/cc66f699-1eb7-400c-9cf7-6a2132f95457) 


----
Need install:
https://pt-br.libreoffice.org/baixe-ja/libreoffice-novo/

 - 1 - Open project "C:\yourLocal\ComponentDinosOffice-OpenOffice\srcPackage\OpenOfficeComponent_install.dproj"
         For delphi 7 use the C:\yourLocal\ComponentDinosOffice-OpenOffice\srcPackag\OpenOfficeComponent_install_Delphi7.dpk
 - 2 - Clean
 - 3 - Build
 - 4 - install
 - 5 - tools -> options -> language -> Delphi -> Library -> Library/ADD:
        "C:\yourLocal\ComponentDinosOffice-OpenOffice\srcPackage\Src" 
        For delphi 7 use the "C:\yourLocal\ComponentDinosOffice-OpenOffice\srcPackage\Src_dx7"
 
 
| Project {versionOficial}   	| Version 	   |
|----------------------------	|------------ |
| 20.1                     	  | tested ✅  | 

Tested Delphi version

| Version  	| Supported 	|
|----------	|-----------	|
| > 12.x   	|    ✅ 	    |
| > 11.x   	|    ✅    	|
| > 10.x   	|    ✅ 	    |
| Seattle  	|    ✅ 	    |
| XE 8     	|    ✅ 	    |
| Delphi 7 	|    ✅    	|

## For Unigui 
 You need add FDGUIxWaitCursor to your serverModule and in your serverModule checked the property AutoCoInitialize
 
![image](https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/d20893ff-e2c0-4e37-a823-33be3175091e)

## For Intraweb
 You need add FDGUIxWaitCursor to your ServerController and in your ServerController change the property ComInitialize for ciMultiThreaded

 ![image](https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/a164b806-ca33-4242-a183-1a62a6882e7b)

 
| Version Tested   	       | Supported 	|
|-------------------------	|-----------	|
| Unigui 1.95.0.1584 	     |    ✅ 	   |
| Intraweb 14    	         |    ✅ 	   |

----
For basic read and write sheet usage, just use the component 
To style the worksheet and documents, use the additional units: uOpenOfficeHelper.pas, uOpenOfficeCollors.pas 


This software is MIT open source!

Wiki: https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/wiki

## See more on Instagram https://www.instagram.com/dinosdev/

### O componente é totalmente free, se ele foi muito útil para você, que tal me pagar um café para incentivar o projeto?

PIX:

<img src="https://github.com/Daniel09Fernandes/ComponentDinosOffice-OpenOffice/assets/29381329/00dcc168-df75-4228-b80d-7262c7b4c478" width="300" height="300">


## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=Daniel09Fernandes/DinosOffice&type=Date)](https://star-history.com/#Daniel09Fernandes/DinosOffice&Date)
