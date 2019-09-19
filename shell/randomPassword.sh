#!/bin/bash
function rand(){
	min=$1       
	max=$(($2-$min+1))       
	num=$(cat /proc/sys/kernel/random/uuid | cksum | awk -F ' ' '{print $1}')       
	echo $(($num%$max+$min))  
}

http_post(){    
	post_data=`curl -X POST -d "$1" http://117.50.43.204:8000/auth/v1/create 2>/dev/null`
	echo $post_data
	return 1
}
rnd=$(rand 100000 999999)
param="username=13888888888&password=$rnd&expire=24"
postResult=$(http_post $param)
address="http://117.50.43.204"
echo " address:$address
 student/organization:13888888888 $rnd
 inspur:wait to complete" | mail -v -s "yunpj PassWord Update" -c onlineeval_inspur@163.com 2681666570@qq.com,1019624637@qq.com
