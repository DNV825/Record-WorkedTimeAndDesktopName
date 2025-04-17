# Record-WorkedTimeAndDesktopName オマケ

## Virtual Desktop モジュールに署名が付与されているか確認する手順

Virtual Desktop モジュールに署名が付与されているか確認したい場合、以下のようにコマンドを実行すればよい。署名は付与されていなかったが、これをもって信頼できないかというとそういうわけでもない。

```powershell
Save-Module -Name VirtualDesktop -Repository PSGallery -Path .\Saved

Get-ChildItem .\Saved\VirtualDesktop\*\*.ps*1 |
    Get-AuthenticodeSignature |
    Select-Object Path, Status, SignerCertificate | Format-Table -Auto

Path                                                                                                                Status SignerCertificate
----                                                                                                                ------ -----------------
C:\<ダウンロードした場所のパス>\Saved\VirtualDesktop\1.5.10\VirtualDesktop.ps1  NotSigned                  
C:\<ダウンロードした場所のパス>\Saved\VirtualDesktop\1.5.10\VirtualDesktop.psd1 NotSigned                  
C:\<ダウンロードした場所のパス>\Saved\VirtualDesktop\1.5.10\VirtualDesktop.psm1 NotSigned  
```
