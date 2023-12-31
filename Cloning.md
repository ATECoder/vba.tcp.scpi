# Cloning

- [Source Code](#Source-Code)
  - [Repositories](#Repositories)
  - [Global Configuration Files](#Global-Configuration-Files)
  - [Packages](#Packages)

<a name="Source-Code"></a>
## Source Code
Clone the repository along with its requisite repositories to their respective relative path.

### Repositories
The repositories listed in [external repositories] are required:

[Core repository]
[Winsock repository] 
[IEEE488 repository] 
[SCPI repository] 

```
git clone https://github.com/ATECoder/vba.core.git
git clone https://github.com/ATECoder/vba.winsock.git
git clone https://github.com/ATECoder/vba.tcp.ieee488.git
git clone https://github.com/ATECoder/vba.tcp.scpi.git
```

Clone the repositories into the following folders (parents of the .git folder):
```
%vba%\core\core
%vba%\iot\winsock
%vba%\iot\tcp.ieee488
%vba%\iot\tcp.scpi
```
where %vba% is the root folder of the VBA libraries, e.g., %my%\lib\vba, and %my%, e.g., c:\my is the overall root folder.

[external repositories]: ExternalReposCommits.csv
[Core repository]: https://github.com/ATECoder/vba.core.git
[Winsock repository]: https://github.com/ATECoder/vba.winsock.git
[IEEE488 repository]: https://github.com/ATECoder/vba.tcp.ieee488.git
[SCPI repository]: https://github.com/ATECoder/vba.tcp.scpi.git
