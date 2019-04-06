# 内容解析规则：

## 1、一行默认为一个命令

## 2、空格分割命令与参数

## 3、逗号分割参数



# 命令如下：

## 去 [其它节点名字] #代表从本节点连接到的节点

## 值 [数字] #代表本节点的初始值

## 算 [运算符号]#+(源+本)、-(源-本)、--(本-源)、*(源*本)、/(源÷本)、//(本÷源)、=(本=源)

## 上限 [数字] #阈值上限，超过将截止于此不进行运算

## 下限 [数字] #阈值下限，低于将截止于此不进行运算

## 常 #出现此字则意味着本节点是常数节点，常数节点的所有运算不会改变自身的值，但可以把自身值和源节点输出值的运算结果传递给下一个节点