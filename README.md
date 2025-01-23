# VersionMapping
A PowerShell tool for file version mapping 


# prerequisite

If you can not run power shell script , you can set the policy ,the temporary solution is to run below command to set policy. 

~~~
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
~~~
Or for long-term solution is to run below command to set policy.

~~~
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
~~~

# install excel model 
If you want to use this script, you must first execute the following command to install the dependent modules.
~~~
Install-Module -Name ImportExcel -Scope CurrentUser
~~~

