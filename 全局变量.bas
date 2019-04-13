Attribute VB_Name = "全局变量"
Public Type 二维坐标
    X As Single
    Y As Single
    End Type
    
Public Type 节点
    名字 As String
    索引 As Long
    内容 As String
    权值 As Double '节点的权值
    大小 As Single
    颜色 As Long
    坐标 As 二维坐标
    去 As String
    去缓存() As Long
    编辑界面偏移 As 二维坐标
    End Type

Public Type 连接
    归 As Long
    去 As Long
    End Type
    
Public 点() As 节点, 线() As 连接, 当前选中点 As Long, 编辑界面装载点 As Long
Public 鼠标位置 As 二维坐标, 启动节点 As String
Public 绘制需求 As Boolean, 节点默认颜色 As Long, 节点默认前缀 As String, 节点默认内容 As String
