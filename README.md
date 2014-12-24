OleDbVCExample
==============

This is a Ole Db example using C++ and SQL Server Express.

It is a console project, but you can use it for your MFC project.

Here are some chinese comments in the project, ignore them. I've provided english comments.

## Directory
***
Files you should care about:
+ `OleDbProject.cpp` main function
+ `OleDbSQL.h OleDbSQL.cpp` OleDb operating class
+ `stdafx.h` declare necessary .h files


## About Database
***
I've used a database named `master` in `SQLEXPRESS` service. Only one table `airport` with two columns(`code`, varchar(10), `name`, varchar(20)) in it. You can change the connection string in `OleDbSQL.cpp` file.

I've referred this article: [url](http://gamebabyrocksun.blog.163.com/blog/static/571534632008101083957499/)(chinese). Some bugs exist in his codes.