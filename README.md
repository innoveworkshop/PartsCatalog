# Parts Catalog

A collection of applications to browse through your components, import orders,
and manage your parts bins.

For more information on each component of this project, please check the README
files inside each of the sub-folders of this repository. There you'll get all
the information about each of the subprojects that make this application
possible, including some nice **screenshots**.


## Database

This program uses an Access 2000 database to store its data and a database
template can be found under the `Database/` folder. The reason to use such an
ancient and crappy database is simply because I wanted this program to be
compatible with old versions of Windows and most importantly be able to be
synced with Windows CE, and for that the only viable database option was Access
2000.


## Building

I'm still working on a [NMake](https://docs.microsoft.com/en-us/cpp/build/reference/nmake-reference?view=msvc-160)
script to automate the build process of the whole project.


## Why Visual Basic 6?

As you can clearly see, this program was written using the ancient VB6 language,
the reason for this is fairly simple: It works on almost any version of Windows,
from the XP to 10 without any issues and it makes programming really fun.

In my humble opinion VB6 is a great way to develop a nice looking professional
GUI without wasting too much time and having lots of fun. If I did this using,
let's say C#, I would waste a lot of time trying to make sure I had everything
neatly wrapped into classes and good programming practices like that, but with
VB6 I can create something nice very fast, and for internal programs like this,
this powerful tool is indispensable and saves me from wasting way too much time
making sure the code is absolutely perfect.


## License

This project is licensed under the **MIT License**.

