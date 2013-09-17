### Export2Excel (Export Mysql/Oracle Data to Excel CSV)

This library is used to generate customized Excel (.XLS ) report from the given SQL queries. It is using an XML file with name config.xml for configurations and another XML file giving SQL queries to generate the report.

####1) config.xml

Here configure the input file name, database connection details and the constants that can be used in the input file.

    <?xml version="1.0"?>
    <config>
	<inputfiles>
		<file>query.xml</file>
	</inputfiles>
	<constants>
		<constant name="$date">28-06-2012</constant>
		<constant name="$sysdate">'27-jun-2012'</constant>
	</constants>
	<database>
		<vendor>mysql</vendor>
		<!-- <path>jdbc:oracle:thin:@192.168.0.1:1521:test</path> -->
		<path>jdbc:mysql://localhost:3306/test</path>
		<username>root</username>
		<password></password>
		<!-- <driver>oracle.jdbc.driver.OracleDriver</driver> -->
		<driver>com.mysql.jdbc.Driver</driver>
	</database>
</config>

####2) query.xml

This file is used to give the input SQL queries and the format for tables to be written in the excel report.

    <?xml version="1.0"?>
    <export>
	<excel name="report.xls">
		<sheet name="test">
			<report name="Test data">
				<query>
					<sql>
						select * from test
					</sql>
				</query>
				<fields>
					<field name="ID" type="column" datatype="number" width="50">id</field>
					<field name="Name" type="column" datatype="text" width="50">name</field>
				</fields>
			</report>
		</sheet>
	</excel>
</export>


####3) Run.bat

This file is used to execute the JAVA program. Just edit the JDK bin path and run the file. The excel file will be generated in the same directory.

	@echo off
	echo Setting Java Path..
	set path="C:\Program Files\Java\jdk1.5.0\bin"
	echo Java Path is set...
	echo Compiling 'KeyMethodMapV1.java'...
	javac -d . lib/KeyMethodMapV1.java
	echo Compiling 'ExportTableV1.java'...
	javac -d . -classpath lib\jdom-1.0.jar;lib\poi-3.0-rc4-20070503.jar;lib\ojdbc14.jar; lib/ExportTableV1.java
	echo Compiling 'ExportExcelV1.java'...
	javac -d . -classpath lib\jdom-1.0.jar;lib\poi-3.0-rc4-20070503.jar;lib\ojdbc14.jar; lib/ExportExcelV1.java
	echo Running 'ExportExcelV1.class'...
	java -classpath lib\jdom-1.0.jar;lib\poi-3.0-rc4-20070503.jar;lib\ojdbc14.jar; lib/ExportExcelV1
	pause
