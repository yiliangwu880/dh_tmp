#!/usr/bin/env bash
#更新bin 文件，重启
#根据配置，筛选一种游戏所有目录更新重启
#开发

REMOTE_IP=192.168.1.203
REMOTE_PSW=123456
BIN_FILE=crapGame			#linux/game_platform/bin 里面的执行文件
TEMPLATE_NAME=crap_game_  	#筛选一种游戏目录
SERVICE_PATH=/home/app/linux/game_platform/run/service/

#远程服务器存放bin的文件夹 列表
#build all_fold_name_list
###################################################
all_fold_name_list=()
FINAL_FOLD_NAME_LIST=$(sshpass -p ${REMOTE_PSW} ssh app@${REMOTE_IP} << eof
	cd ${SERVICE_PATH};
	ls -d ${TEMPLATE_NAME}*
eof
)
for FINAL_FOLD_NAME in ${FINAL_FOLD_NAME_LIST}
do
	all_fold_name_list[${#all_fold_name_list[*]}]=${SERVICE_PATH}${FINAL_FOLD_NAME}
done

###################################################


#更新stop 更bin, start
#$1 fold_name list
function UpdateAll()
{
   fold_name_list=("${!1}")
	
   for v in ${fold_name_list[@]} ;do
		echo update: ${v}
   done
   echo =========stop
   for v in ${fold_name_list[@]} ;do
		sshpass -p ${REMOTE_PSW} ssh app@${REMOTE_IP}  "cd ${v}; ./stop.sh"
   done
   sleep 3
   echo =========copy
   for v in ${fold_name_list[@]} ;do
		sshpass -p ${REMOTE_PSW} scp ../../../../bin/${BIN_FILE} app@${REMOTE_IP}:${v}
   done
   sleep 1
   echo =========start
   for v in ${fold_name_list[@]} ;do
		sshpass -p ${REMOTE_PSW} ssh app@${REMOTE_IP}  "cd ${v}; ./start.sh"
   done
}

UpdateAll all_fold_name_list[@]
echo ====end====  