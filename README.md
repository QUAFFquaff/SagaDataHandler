7/4 新任务：数据处理
将txt文件中的数据筛选，分类，统计到excel中；
尝试使用csv，太丑，使用xlwt，但文件太大不支持，最终使用xlsxwriter包；
生成 五个excel，关于五个同设备类型的相关点位编码信息，按deviceUID分sheet，按时间升序，有几个文件还挺大，也得处理一会，受限于python 的I/O，暂时没有优化方案。（pandas不太方便）