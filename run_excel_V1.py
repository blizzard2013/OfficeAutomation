import xlwings as xw 
import datetime


# 实例化设置
app = xw.App(visible=True,add_book=False)
# app = xw.App(visible=False,add_book=False)
# excel报警设置，默认true
# app.display_alerts = False
# 屏幕刷新设置,默认true
# app.screen_updating=False

# 1 定义筛选条件
# # 统计关键字
checktime1='2019-3-1'			# 指定日期
checktime2=15					# 15天
checktime3=183				# 6个月

# # 部门关键字
DN1 = '底盘'
DN2 = '吉利'
DN3 = '新能源'
DN4 = '长兴'

# # 粘贴位置定义
pass

# # condition条件文件位置
# cd_path=os.getcwd()		# 当前路径
# condition条件位置
c_start_cow=5		# 起始行数
cd_state_c='c'		# 状态
cd_res_c='f'		# 责任人
cd_dep_c='g'		# 部门:通配符匹配问题？
cd_sys_c='r'		# 系统
cd_mould='B2:Q37'	# 统计表模板

# # source源数据文件位置
# s_path=input("请输入文件路径")
# source源数据位置
s_start_cow=5		# 起始行数
s_state_c='c'		# 状态-------ok
s_res_c='f'			# 责任人-----ok
s_dep_c='g'			# 部门XXXXXXXX
s_sys_c='r'			# 系统
s_c_data_c='bf'		# 创建
s_a_data_c='et'		# 指派



# 2 另存筛选后工作簿
# 打开模板工作簿
wb = app.books.open('★GPIS问题筛选_AtoB_python_模板_V2.xlsx')
# 生成date及str格式当前日期
now = datetime.datetime.now().date()
nowstr = now.strftime("%Y-%m-%d")
# 另存为新工作簿2
# nwb.save(f'设计实车验证问题-未关闭-PSS100-{nowstr}.xlsx')
wb.save(f'设计实车验证问题-未关闭-PSS100-{nowstr}.xlsx')

# 2 读取筛选条件
# 指定工作表			# 【条件来源】
conditionsht = wb.sheets[1]
# 读取条件
cd_state=conditionsht.range(f'{cd_state_c}{c_start_cow}').expand('down').value
cd_res=conditionsht.range(f'{cd_res_c}{c_start_cow}').expand('down').value
# cd_dep=conditionsht.range(f'{cd_dep_c}{c_start_cow}').expand('down').value			# 不用通配符，USELESS
cd_sys=conditionsht.range(f'{cd_sys_c}{c_start_cow}').expand('down').value
# print(cd_state)
# print(cd_res)
# print(cd_dep)
# print(cd_sys)

# 添加并指定工作表—			# 【问题保存】
savesht = wb.sheets.add(name=None,before=None,after=conditionsht)

# 删除条件工作表
conditionsht.delete()

# 3 打开源数据，筛选符合条件的行并复制粘贴到新文件中

# 打开数据源工作簿
swb = app.books.open('设计实车验证问题.xls')
# 指定工作表—			# 【问题来源】
sourcesht = swb.sheets[0]

# 粘贴标题行
# title=sourcesht.range(f'{s_start_cow-1}:{s_start_cow-1}').value					# 取整行值，可能有隐患，其他办法？？？？？？？？
title=sourcesht.range(f'a{s_start_cow-1}').expand('right').value
savesht.range('a7').value=title
# savesht.range('a7').row_height=13.5
# input()

# 获取源数据最大范围
maxrange = sourcesht.range(f'a{s_start_cow-1}').expand('table')
maxrows = maxrange.rows.count
# maxcoloums = maxrange.columns.count
maxcoloums = 'fl'
# print(maxrows)			#380---383=380+5-2
# print(maxcoloums)		#168

# 筛选符合要求的问题并依序粘贴到指定位置
i=8					# 符合名单的问题起始位置
j=47				# 不符合名单但符合专业的问题起始位置
for r in  list(range(s_start_cow,s_start_cow+maxrows)):
	if cd_state.count(sourcesht.range(f'{s_state_c}{r}').value):		# 状态判定		包含确认：列表count，字符串find，in
		if cd_res.count(sourcesht.range(f'{s_res_c}{r}').value):		# 责任人判定		包含确认：列表count，字符串find，in
			record=sourcesht.range(f'{r}:{r}').value					# 取整行值，有隐患：取值不完整+无法带格式，其他办法？？？？？？？？
			# record=sourcesht.range(f'a{r}:{maxcoloums}{r}').value
			record[2]="'" + record[2]							# 状态数值+'转换为str，以免写入时格式变化为日期
			# if record[7]!='':
				# record[7]=datetime.datetime.strptime(record[7],"%Y-%m-%d %H:%M:%S").date()			# str日期转date——————也可在后面标记颜色时调整
			savesht.range(f'a{i}').value=record
			savesht.range(f'a{i}').row_height=13.5
			i+=1	
		elif DN1 in sourcesht.range(f'{s_dep_c}{r}').value \
			and ( DN2 in sourcesht.range(f'{s_dep_c}{r}').value \
			or DN3 in sourcesht.range(f'{s_dep_c}{r}').value \
			or DN4 in sourcesht.range(f'{s_dep_c}{r}').value ):
			if cd_sys.count(sourcesht.range(f'{s_sys_c}{r}').value):	# 专业判定		包含确认：列表count，字符串find，in
				record=sourcesht.range(f'{r}:{r}').value				# 取整行值，有隐患：取值不完整+无法带格式，其他办法？？？？？？？？
				record[2]="'" + record[2]						# 状态数值+'转换为str，以免写入时格式变化为日期
				savesht.range(f'a{j}').value=record
				savesht.range(f'a{j}').row_height=13.5
				j+=1

hide=savesht.range('H1')
hide.column_width=14.7
hide=savesht.range('M1')
hide.column_width=14.7
hide=savesht.range('BF1')
hide.column_width=14.7
hide=savesht.range('ET1')
hide.column_width=14.7

swb.close()

# # # # # # # # # # # # # 4 格式调整(隐藏&组合)==================================================================

hide=savesht.range('I1:K1')
hide.column_width=0
hide=savesht.range('O1:Q1')
hide.column_width=0
hide=savesht.range('S1:AT1')
hide.column_width=0
hide=savesht.range('AV1:BE1')
hide.column_width=0
hide=savesht.range('BG1:DC1')
hide.column_width=0
hide=savesht.range('DF1:DF1')
hide.column_width=0
hide=savesht.range('DH1:ES1')
hide.column_width=0

pass

# 5 建立新工作簿2，插入工作表，并复制粘贴统计表

# 指定粘贴工作表
countsht = wb.sheets[0]																		#	!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
countsht.name='统计'

# # # # # # # # # # 插入工作表
# # # # # # # # # nwb.sheets.add(name='统计',before=savesht)
# # # # # # # # # # 复制粘贴统计表
# # # # # # # # # matrix=countsht.range(cd_mould).value
# # # # # # # # # # matrix=countsht.range(cd_mould).api.Copy
# # # # # # # # # countsht=nwb.sheets[0]
# # # # # # # # # countsht.range('b2').value=matrix
# # # # # # # # # # countsht.range('b2').api.Paste=matrix
# # # # # # # # # # countsht.range(cd_mould).api.Copy.countsht.range('b2')
# # # # # # # # # # countsht.range('b2')=matrix
# # # # # # # # # # countsht.range('b2')=countsht.range(cd_mould).api.Copy
# # # # # # # # # # # 带格式复制粘贴待完善===================================

# # 调整统计表格式
# countsht.autofit('c')

# # # # # # # # # # 颜色边框？==================================================================



# # # # # # # # # # # 6 统计数据：公式、累加、标记颜色
# 添加公式
countsht.range('d3:d30').formula='=VLOOKUP(CONCATENATE("*",$C3,"*"),Sheet1!$A:$M,6,FALSE)'
countsht.range('e3:e30').formula='=COUNTIF(Sheet1!$A:$A,CONCATENATE("*",$C3,"*"))'
countsht.range('f3:f30').formula='=COUNTIFS(Sheet1!$A:$A,CONCATENATE("*",$C3,"*"),Sheet1!$C:$C,"0/4")'
countsht.range('g3:g30').formula='=COUNTIFS(Sheet1!$A:$A,CONCATENATE("*",$C3,"*"),Sheet1!$C:$C,"1/4")'
countsht.range('e31:k31').formula='=SUM(E3:E30)'
# # 表格区域自带公式，公式输入可省略
# countsht.range('m33').formula='=E31'
# countsht.range('n33').formula='=H31'
# countsht.range('o33').formula='=F31+G31'
# countsht.range('p33').formula='=J31'
# countsht.range('q33').formula='=K31'
# 计算符合条件的行数&标记颜色
# # 清空统计区域内容避免干扰
countsht.range('h3:k30').value = 0
# # 确认需要统计的区域最大行数——————TBD
maxrows=savesht.range('A65536').end('up').row				# A列最后一行有内容的行序号
# maxrange = sourcesht.range(f'a{s_start_cow-1}').expand('table') # 只能统计最大连续行-------断开行/最后一行有内容的部分怎么统计？
# maxrows = maxrange.rows.count
# VBA # # Sheets("sheet1").Select
# VBA # # imax = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row——最大行数
# # 根据区域确定的起始位置
# i=8					# 符合名单的问题起始位置
# j=37				# 不符合名单但符合专业的问题起始位置

# 根据条件累加并标记颜色:红(255,0,0)、黄(255, 255, 0)、橙(255, 192, 0)、绿(0, 176, 80)
for r in  list(range(8,47)):		# 问题区域
	print(f'BF{r}')
	print(savesht.range(f'BF{r}').value)
	print(type(savesht.range(f'BF{r}').value))

	# 橙色标记：超6个月标记：创建日期BF
	if savesht.range(f'BF{r}').value==None:
		pass
	elif now > savesht.range(f'BF{r}').value.date() + datetime.timedelta(days = checktime3):			# 条件筛选:创建时间距今超过checktime3=183d=6m
		savesht.range(f'BF{r}').color=(255, 77, 0)			# 颜色标记:橙(255, 192, 0)
		# savesht.range(f'A{r}:L{r}').color=(255, 192, 0)			# 颜色标记:橙(255, 192, 0)			# 不重要，暂时无需重点标记
		for j in range(3,31):										# 整个统计表区域
			if countsht.range(f'C{j}').value in savesht.range(f'A{r}').value:		# 包含确认：列表count，字符串find，in
				countsht.range(f'K{j}').value += 1				# 累加

	# 橙色标记：录入时间标记:录入时间BF
	if savesht.range(f'BF{r}').value==None:
		pass
	elif savesht.range(f'BF{r}').value.date() < datetime.datetime.strptime(checktime1, "%Y-%m-%d").date():			# 条件筛选:(未关闭问题)录入时间小于checktime1=2010-3-160d
		savesht.range(f'BF{r}').color=(255, 192, 0)			# 颜色标记:橙(255, 192, 0)
		# savesht.range(f'A{r}:L{r}').color=(255, 192, 0)			# 颜色标记:橙(255, 192, 0)			# 不重要，暂时无需重点标记
		for j in range(3,31):										# 整个统计表区域
			if countsht.range(f'C{j}').value in savesht.range(f'A{r}').value:		# 包含确认：列表count，字符串find，in
				countsht.range(f'I{j}').value += 1				# 累加

	# 黄色标记：超15天标记:指派时间ET，录入时间BF
	if savesht.range(f'ET{r}').value == None:
		starttime = savesht.range(f'BF{r}').value
	else:
		starttime = savesht.range(f'ET{r}').value
	if (savesht.range(f'C{r}').value == "0/4" or savesht.range(f'C{r}').value == "1/4") and now > starttime.date() + datetime.timedelta(days = checktime2):			# 条件筛选:0/4或1/4状态且指派时间（或提出时间）超过checktime2=15d
		savesht.range(f'ET{r}').color=(255, 255, 0)			# 颜色标记:黄(255, 255, 0)
		savesht.range(f'A{r}:R{r}').color=(255, 255, 0)			# 颜色标记:黄(255, 255, 0)
		for j in range(3,31):										# 整个统计表区域
			if countsht.range(f'C{j}').value in savesht.range(f'A{r}').value:		# 包含确认：列表count，字符串find，in
				countsht.range(f'J{j}').value += 1				# 累加

	# 红色标记:当前责任人L,当前步骤M,
	if savesht.range(f'M{r}').value==None:
		pass
	elif savesht.range(f'M{r}').value.date() < now:			# 条件筛选:当前步骤完成日期小于当前日期
		savesht.range(f'L{r}:M{r}').color = (255,0,0)			# 颜色标记:红(255,0,0)
		savesht.range(f'A{r}:F{r}').color=(255,0,0)			# 颜色标记:红(255,0,0)
		for j in range(3,31):										# 整个统计表区域
			if countsht.range(f'C{j}').value in savesht.range(f'A{r}').value:		# 包含确认：列表count，字符串find，in
				countsht.range(f'H{j}').value += 1				# 累加

	# 红色标记:计划关闭H(NONE)
	if savesht.range(f'H{r}').value==None:
		pass
	elif savesht.range(f'H{r}').value.date() < now:			# 条件筛选:关闭日期小于当前日期
		savesht.range(f'H{r}').color = (255,0,0)			# 颜色标记:红(255,0,0)
		savesht.range(f'A{r}:F{r}').color=(255,0,0)			# 颜色标记:红(255,0,0)
		for j in range(3,31):										# 整个统计表区域
			if countsht.range(f'C{j}').value in savesht.range(f'A{r}').value:		# 包含确认：列表count，字符串find，in
				countsht.range(f'H{j}').value += 1				# 累加

	# # 红色标记:当前责任人L,当前步骤M,计划关闭H(NONE)
	# if savesht.range(f'L{r}').color == (255,0,0):			# 条件筛选:颜色 红
		# savesht.range(f'A{r}:F{r}').color=(255,0,0)			# 颜色标记:红(255,0,0)
		# for j in range(3,31):										# 整个统计表区域
			# if countsht.range(f'C{j}').value in savesht.range(f'A{r}').value:		# 包含确认：列表count，字符串find，in
				# countsht.range(f'H{j}').value += 1				# 累加

# 7 整理(删除空行)，保存，关闭
# 获取统计表区域最大行数
maxrange2 = countsht.range('b2').expand('table')
maxrows2 = maxrange2.rows.count
# 删除符合要求的行(e列=0)
for r2 in  reversed(list(range(2,2+maxrows2))):
	# print(countsht.range(f'e{r2}').value)
	if countsht.range(f'e{r2}').value==0:
		countsht.range(f'e{r2}').api.EntireRow.Delete()

# 保存并关闭文件
wb.save()
wb.close()


# 退出excel
app.quit()

