# VBA-kantankanri



## これは何？

マクロ付のExcelブックを、マクロなしExcelブックとマクロに分けるツールです。分けたものを再度くっつけることもできます。



## 何で必要なの？

gitやSVNでバージョン管理を容易にするためです。



## 使い方は?

次のような構成を想定します。

```
任意のフォルダ ┳ bin ━ 対象.xlsm
			 ┣src
　			┣template
　　　　　　　 ┗kantan-kanri.ps1
```

以下のコマンドでマクロなしExcelブックとマクロが作成されます。

``` 
> powershell (任意のフォルダ)\kantan-kanri.ps1 "対象.xlsm","export"
```



