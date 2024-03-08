#COMMAND_DIR=$(cd "$(dirname "$0")";pwd)
#echo $COMMAND_DIR
# 方法一：目录有中文时会乱码，脚本会失败

COMMAND_PATH=$0
echo "COMMAND_PATH" $COMMAND_PATH
COMMAND_DIR=${COMMAND_PATH:0:${#COMMAND_PATH}-18}
# 获取文件所处的绝对路径
# 18这个值是 excel2word.command的长度，因为遇到有中文乱码的问题，使用这个方法来保证可以正常获取到对应的（带空格及带中文）文件路径，未来可以优化
echo "COMMAND_DIR" $COMMAND_DIR
# 方法二：兼容中文

cd "$COMMAND_DIR"
python3 excel2word.py "${COMMAND_DIR}"
