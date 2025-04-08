"""
name : 小米商城做任务
cron : 1 9 * * *
脚本兼容 : 青龙
更新时间 : 20241120
环境变量名：mi
环境变量值：cookie
备注：抓小米商城app的包。有黑号风险。若提示更新，请将任务列表响应发到https://github.com/imoki/mi/仓库的issues中。
    提交到imoki/mi仓库issues的抓包内容为：抓取小米商城app的https://shop-api.retail.mi.com/mtop/navi/venue/batch包的响应体，将响应全部发送到issues中，请务必参考issues模板发送。
    如何判断提交的响应体格式是否正确：响应中含有doneInSingleAward。

在github中发送issues:
1. 提交的标题格式如下：（标题格式：抓包时间+任务列表响应）
20241117任务列表响应
2. 提交的内容格式如下：（响应中含有doneInSingleAward）
{"traceId":"","code":0,"attachments":{"timestamp":""},"data":{"result_list":[{"amountTotal":760,"jumpUrl":"","tipList":["完成每日签到赚米金","积攒米金兑超级好物"],"waitArrival":0,"waitReceiveList":[],"waitReceiveTotal":0},{"components":[{"actId":"6706c0695404a23dfb5b2cab","taskId":"6706c0695243011f230d465d","taskName":"米金签到","taskDesc":"","taskType":200,"finishedNumber":0},{"actId":"6706c0695404a23dfb5b2cab","taskId":"6706edf30344c966c5b46681","taskName":"来会员中心 领专属好券","taskDesc":"浏览10秒+10米金","taskType":200,"finishedNumber":0},{"actId":"6706c0695404a23dfb5b2cab","taskId":"670f6baf1d65ee598c4fc39d","taskName":"新客专享福利","taskDesc":"浏览10秒+10米金","taskType":200,"finishedNumber":0},{"actId":"6706c0695404a23dfb5b2cab","taskId":"670fa2ba57e9a97a89265b63","taskName":"逛手机频道 选心动手机","taskDesc":"浏览10秒+10米金","taskType":200,"finishedNumber":0},{"actId":"6706c0695404a23dfb5b2cab","taskId":"6720abdc4015d32aaaa400bc","taskName":"Xiaomi 15系列 新品手机","taskDesc":"浏览10秒+10米金","taskType":200,"finishedNumber":0}]},{"components":[{"actId":"","taskId":"","taskName":"米金抽奖","taskDesc":"","taskType":128,"finishedNumber":0,"totalNumber":20,"startTime":1720000000000,"endTime":1730000000000,"status":2,"scores":0,"singleCostScores":30,"costType":2,"cycleTime":0,"taskRefreshWay":0,"upperLimit":20,"canDo":true}]}]},"message":"ok"}

"""

#!/usr/bin/python3
import lzma, base64
exec(lzma.decompress(base64.b64decode('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')))