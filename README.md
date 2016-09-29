ExcelUtil
==
####The project is based on POI .If you want to konw more about it,just click the right link :[POI](http://poi.apache.org/)   
##How to use the ExcelUtil ? don't worry,You just need to do the following steps.
Step 1:   
```java
allprojects {
    repositories {
        jcenter()
        maven { url "https://jitpack.io" }
    }
}
```

Step 2:
```java
 compile 'com.github.sqyNick:CExcelUtil:1.0.7'
```
##Then you can use the ExcelUtil just like below.
```java
        ArrayList<ArrayList<ArrayList<Object>>> excel = ExcelUtil.readExcel(new File("your excel path"));
        Object cell = excel.get(0).get(0).get(0);//Get zeroth Sheet zeroth rows and 0 columns
```
