# Order Importer

The easiest way to import a whole distributor order into a
[PartsCatalog](https://github.com/innoveworkshop/PartsCatalog) database.

![Screenshot](/Screenshots/2020-11-18.png)

## Supported Distributors

Currently the only supported distributor is
[Farnell Portugal](https://pt.farnell.com/), although I think other Farnell
branches might also be supported, but I haven't tried them yet.

Another thing to take into consideration is that this program is supposed to be
used with the CSV file that is exported from the Order Status view, not a
shopping cart.


## Building

If you want to build this program or contribute to it make sure that you have
the appropriate environment set up. You'll need a copy of Visual Basic 6.0,
which you can eaily download from [WinWorld](https://winworldpc.com/product/microsoft-visual-stu/60),
if you plan on installing it under any modern version of Windows, make sure to
search for a tutorial first before you litter your system with old, unuseable,
crap. After that's all done you'll need to download a copy of the
[PartsCatalog](https://github.com/innoveworkshop/PartsCatalog) project and have
a folder structure that looks somewhat like this:

    \<Dev Folder>\OrderImporter\
	\<Dev Folder>\PartsCatalog\

The name and relative path of the project folders is important because
OrderImporter requires some components from the PartsCatalog project that are
shared, and VB6 uses relative paths for that, so the names and locations are
important.


## Why Visual Basic 6?

I've used the dreaded VB6 for this project for two main reasons: First of all
nostalgia, I've written this partially on a Windows XP VM, but most of it was
written on a Windows 2000 laptop. I just enjoy using old gear. The second reason
was because I wanted to have this working on any old operating system I might be
inclined to use in the future. I could've written it using
[Lazarus](https://www.lazarus-ide.org/), but I wasn't feeling particularly
interested in Pascal the day I started the project, so yeah, deal with it.


## License

This project is licensed under the **MIT License**.

