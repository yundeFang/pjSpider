#!/bin/bash
function rand(){
	min=$1       
	max=$(($2-$min+1))       
	num=$(cat /proc/sys/kernel/random/uuid | cksum | awk -F ' ' '{print $1}')       
	echo $(($num%$max+$min))  
}

http_post(){    
	post_data=`curl -X POST -d "$1" http://117.50.43.204:8000/auth/v1/resetpassword 2>/dev/null`
	return 1
}
rnd=$(rand 100000 999999)
old_pas=$(cat /opt/old.pas)
param1="username=666666&scope=1&old_password=$old_pas&new_password=$rnd"
param2="username=999999&scope=2&old_password=$old_pas&new_password=$rnd"
param3="username=201909032&scope=3&old_password=$old_pas&new_password=$rnd"
postResult=$(http_post $param1)
postResult=$(http_post $param2)
postResult=$(http_post $param3)
echo $rnd > /opt/old.pas

address="http://117.50.43.204"
echo " address:$address
 inspur:       666666    $rnd
 organization: 999999    $rnd
 student:      201909032 $rnd " | mail -v -s "yunpj PassWord Update" -c onlineeval_inspur@163.com onlineeval_inspur@163.com