Attribute VB_Name = "全局变量"
Public Type 阀值
    '当下限>上限时，所有信息都将通过该节点
    下限 As Double '包含该值
    上限 As Double
    低限导向 As String
    超限导向 As String
    输出下限 As Double
    输出上限 As Double
    End Type

Public Type 不应周期
    位 As Long '每触发一次信号接入，如果为0，则触发运算且变为不应期长度，不为0则减1不触发运算
    期 As Long
    常 As Long
    常位 As Long
    去 As String
    End Type

Public Type 遗忘周期
    原值 As Double
    期 As Long
    位 As Long
    End Type

Public Type 锁
    次 As Long '能够触发的最大有效次
    End Type

Public Type 二维坐标
    X As Single
    Y As Single
    End Type
    
Public Type 节点
    名字 As String
    索引 As Long
    激活数 As Long
    内容 As String '存在节点内的规则
    权值 As Double '节点的权值
    流值 As Double '仅常量使用
    大小 As Single
    遗忘 As 遗忘周期
    常量 As Boolean  '如果该节点是常量则不会被改变值，但仍然可以传递运算值
    阈值 As 阀值
    不应 As 不应周期
    自锁 As 锁
    运算 As String ' +-*/ 四种基本运算
    颜色 As Long
    信息流 As String '记载着信息经过的所有节点名字
    坐标 As 二维坐标
    去 As String
    编辑界面偏移 As 二维坐标
    End Type

Public Type 连接
    归 As Long
    去 As Long
    End Type
Public 点() As 节点, 线() As 连接, 当前选中点 As Long, 编辑界面装载点 As Long
Public 鼠标位置 As 二维坐标, 启动节点 As String
Public 绘制需求 As Boolean, 节点默认颜色 As Long, 节点默认前缀 As String, 节点默认内容 As String
