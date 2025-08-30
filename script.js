// 全局存储当前订单数据
        let currentOrders = [];
        let urgentOrdersList = []; // 存储当前加急订单
        let currentFileName = ""; // 存储当前文件名
        let currentFileExtension = ""; // 存储当前文件扩展名

        // 状态颜色映射
        const statusColors = {
            '已完成': '#4BC0C0',
            '待支付': '#FF6384',
            '待发货': '#FFCE56',
            '取消': '#CCCCCC',
            '退款中': '#9966FF'
        };

        // 初始化页面
        document.addEventListener('DOMContentLoaded', function() {
            const selectFileBtn = document.getElementById('selectFileBtn');
            const fileInput = document.getElementById('fileInput');
            const dropArea = document.getElementById('dropArea');
            const uploadContent = document.getElementById('uploadContent');
            const processingStatus = document.getElementById('processingStatus');

            // 文件选择按钮
            selectFileBtn.addEventListener('click', function() {
                fileInput.click();
            });

            // 文件输入变化
            fileInput.addEventListener('change', function(e) {
                if (e.target.files.length > 0) {
                    processFile(e.target.files[0]);
                }
            });

            // 拖放功能
            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                dropArea.addEventListener(eventName, preventDefaults, false);
                document.addEventListener(eventName, preventDefaults, false);
            });

            function preventDefaults(e) {
                e.preventDefault();
                e.stopPropagation();
            }

            ['dragenter', 'dragover'].forEach(eventName => {
                dropArea.addEventListener(eventName, function() {
                    dropArea.classList.add('drag-over');
                }, false);
            });

            ['dragleave', 'drop'].forEach(eventName => {
                dropArea.addEventListener(eventName, function() {
                    dropArea.classList.remove('drag-over');
                }, false);
            });

            dropArea.addEventListener('drop', function(e) {
                const dt = e.dataTransfer;
                const files = dt.files;
                if (files.length) {
                    processFile(files[0]);
                }
            }, false);

            // 导出按钮 - 修复问题：只导出加急订单
            document.getElementById('exportBtn').addEventListener('click', function() {
                if (urgentOrdersList.length === 0) {
                    alert('当前没有加急订单可导出');
                    return;
                }

                try {
                    // 创建工作簿
                    const wb = XLSX.utils.book_new();

                    // 创建一个新的数组，排除内部属性
                    const exportData = urgentOrdersList.map(order => {
                        const { daysPending, ...rest } = order;
                        return rest;
                    });

                    const ws = XLSX.utils.json_to_sheet(exportData);

                    // 添加工作表到工作簿
                    XLSX.utils.book_append_sheet(wb, ws, "加急订单数据");

                    // 导出Excel
                    XLSX.writeFile(wb, `加急订单_${currentFileName}.xlsx`);

                    // 显示导出成功通知
                    showNotification('成功导出加急订单数据', 'success');
                } catch (error) {
                    console.error('导出失败:', error);
                    showNotification('导出失败: ' + error.message, 'danger');
                }
            });
        });

        // 显示通知函数
        function showNotification(message, type) {
            // 如果已经有通知存在，先移除
            const existingAlert = document.querySelector('.notification-alert');
            if (existingAlert) {
                existingAlert.remove();
            }

            const alertBox = document.createElement('div');
            alertBox.className = `alert alert-${type} alert-dismissible fade show notification-alert`;
            alertBox.style.top = '20px';
            alertBox.style.right = '20px';
            alertBox.style.zIndex = '1050';
            alertBox.innerHTML = `
                ${message}
                <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
            `;

            document.body.appendChild(alertBox);

            // 3秒后自动关闭
            setTimeout(() => {
                const bsAlert = new bootstrap.Alert(alertBox);
                bsAlert.close();
            }, 3000);
        }

        // 处理上传的文件 - 修复日期处理逻辑
        function processFile(file) {
            const dropArea = document.getElementById('dropArea');
            const uploadContent = document.getElementById('uploadContent');
            const processingStatus = document.getElementById('processingStatus');

            // 存储文件名和扩展名
            const fileName = file.name;
            const nameExtension = fileName.split('.');
            currentFileName = nameExtension[0];
            currentFileExtension = nameExtension.length > 1 ? nameExtension.pop() : '';

            // 更新当前数据区域的文件名显示
            document.getElementById('fileDisplay').innerHTML = `
                <i class="bi bi-file-earmark-${currentFileExtension === 'xlsx' ? 'excel' : 'text'}"></i>
                <span>${fileName}</span>
            `;

            // 显示处理状态
            uploadContent.style.display = 'none';
            processingStatus.style.display = 'block';

            const reader = new FileReader();

            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    // 解析Excel文件
                    const workbook = XLSX.read(data, {type: 'array'});

                    // 获取第一个工作表
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];

                    // 转换为JSON
                    const orders = XLSX.utils.sheet_to_json(worksheet);

                    // 简单清洗数据
                    const cleanedOrders = orders.map(order => {
                        // 日期转换 - 修复日期格式处理问题
                        if (order['订单日期']) {
                            let parsedDate = null;
                            const dateStr = order['订单日期'];

                            // 尝试解析为Date对象
                            if (typeof dateStr === 'string') {
                                // 尝试标准日期格式
                                if (dateStr.match(/\d{4}-\d{1,2}-\d{1,2}/)) {
                                    const [year, month, day] = dateStr.split('-');
                                    parsedDate = new Date(year, month - 1, day);
                                }
                                // 尝试斜杠格式
                                else if (dateStr.match(/\d{4}\/\d{1,2}\/\d{1,2}/)) {
                                    const [year, month, day] = dateStr.split('/');
                                    parsedDate = new Date(year, month - 1, day);
                                }
                                // 尝试点分隔格式
                                else if (dateStr.match(/\d{4}\.\d{1,2}\.\d{1,2}/)) {
                                    const [year, month, day] = dateStr.split('.');
                                    parsedDate = new Date(year, month - 1, day);
                                }
                                // 尝试中文日期格式
                                else if (dateStr.match(/\d{4}年\d{1,2}月\d{1,2}日/)) {
                                    const [year, month, day] = dateStr.split(/[年月日]/).filter(Boolean);
                                    parsedDate = new Date(year, month - 1, day);
                                }
                            }
                            // 尝试解析为Excel日期序列号
                            else if (typeof dateStr === 'number') {
                                // Excel日期序列号转换（自1900年1月1日起的天数）
                                const baseDate = new Date(1900, 0, 1);
                                baseDate.setDate(baseDate.getDate() + dateStr - 2); // Excel日期系统的偏差
                                parsedDate = baseDate;
                            }

                            // 如果解析成功，格式化为 YYYY-MM-DD
                            if (parsedDate instanceof Date && !isNaN(parsedDate.getTime())) {
                                const year = parsedDate.getFullYear();
                                const month = String(parsedDate.getMonth() + 1).padStart(2, '0');
                                const day = String(parsedDate.getDate()).padStart(2, '0');
                                order['订单日期'] = `${year}-${month}-${day}`;
                                // 添加内部属性用于日期计算
                                order['_parsedDate'] = parsedDate;
                            } else {
                                // 无法解析，保留原值
                                order['订单日期'] = dateStr;
                            }
                        }

                        // 确保金额字段是数字
                        if (order['订单金额'] && typeof order['订单金额'] === 'string') {
                            order['订单金额'] = parseFloat(order['订单金额'].replace(/[^\d.-]/g, ''));
                        }

                        // 确保其他数字字段是数字
                        if (order['商品数量'] && typeof order['商品数量'] === 'string') {
                            order['商品数量'] = parseInt(order['商品数量'], 10);
                        }

                        // 处理订单状态
                        if (order['订单状态']) {
                            const status = order['订单状态'];
                            // 标准化状态名称
                            if (status.includes('待付') || status.includes('未付')) {
                                order['订单状态'] = '待支付';
                            } else if (status.includes('待发') || status.includes('未发')) {
                                order['订单状态'] = '待发货';
                            } else if (status.includes('完成') || status.includes('已完')) {
                                order['订单状态'] = '已完成';
                            } else if (status.includes('取消')) {
                                order['订单状态'] = '取消';
                            } else if (status.includes('退款')) {
                                order['订单状态'] = '退款中';
                            }
                        }

                        return order;
                    });

                    // 更新全局数据
                    currentOrders = cleanedOrders;

                    // 更新UI
                    updateDashboard(cleanedOrders);

                    // 关闭模态框
                    bootstrap.Modal.getInstance(document.getElementById('uploadModal')).hide();

                    // 显示导入成功通知
                    showNotification(`成功导入文件: ${fileName}`, 'success');

                } catch (error) {
                    console.error('文件处理错误:', error);
                    showNotification('处理文件时出错: ' + error.message, 'danger');
                } finally {
                    // 恢复上传UI状态
                    uploadContent.style.display = 'block';
                    processingStatus.style.display = 'none';
                }
            };

            reader.onerror = function() {
                showNotification('读取文件时出错', 'danger');
                uploadContent.style.display = 'block';
                processingStatus.style.display = 'none';
            };

            reader.readAsArrayBuffer(file);
        }

        // 更新仪表盘数据 - 修复处理时长计算
        function updateDashboard(orders) {
            if (orders.length === 0) {
                return;
            }

            // 隐藏空状态
            document.getElementById('emptyState').style.display = 'none';

            // 显示数据和图表
            document.getElementById('statsContainer').classList.add('row');
            document.getElementById('chartsContainer').style.display = 'flex';
            document.getElementById('tablesContainer').style.display = 'block';

            // 计算统计信息
            const totalOrders = orders.length;

            // 总销售额计算
            const totalSales = orders.reduce((sum, order) =>
                sum + (parseFloat(order['订单金额']) || 0), 0).toFixed(2);

            // 获取所有国家并计数
            const countryMap = {};
            orders.forEach(order => {
                const country = order['客户国家'] || '未知';
                countryMap[country] = (countryMap[country] || 0) + 1;
            });
            const countries = Object.keys(countryMap).length;

            // 计算加急订单 - 修复处理时长计算
            const today = new Date();
            today.setHours(0, 0, 0, 0); // 设置为今天的开始时间

            urgentOrdersList = orders.filter(order => {
                if (!order['订单日期'] || !order['_parsedDate']) return false;

                // 使用已解析的日期对象
                const orderDate = new Date(order['_parsedDate']);
                orderDate.setHours(0, 0, 0, 0); // 设置为订单日期的开始时间

                // 计算天数差（精确天数）
                const timeDiff = today - orderDate;
                const daysPending = Math.floor(timeDiff / (1000 * 60 * 60 * 24));
                order.daysPending = daysPending;

                // 标记超过3天未处理的订单为加急订单
                return daysPending > 3 && order['订单状态'] === '待发货';
            });

            const urgentCount = urgentOrdersList.length;

            // 更新统计卡片
            document.getElementById('statsContainer').innerHTML = `
                <div class="col-md-3">
                    <div class="card stat-card">
                        <div class="value">${totalOrders}</div>
                        <div class="label">总订单数</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card stat-card">
                        <div class="value">¥${totalSales}</div>
                        <div class="label">总销售额</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card stat-card">
                        <div class="value">${urgentCount}</div>
                        <div class="label">加急订单</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card stat-card">
                        <div class="value">${countries}</div>
                        <div class="label">国家/地区</div>
                    </div>
                </div>
            `;

            // 更新加急订单表格
            const urgentBadge = document.getElementById('urgentBadge');
            const urgentTableBody = document.querySelector('#urgentTable tbody');
            const urgentNoData = document.getElementById('urgentNoData');

            urgentBadge.textContent = `${urgentCount} 个待处理`;
            urgentTableBody.innerHTML = '';

            if (urgentCount === 0) {
                urgentNoData.style.display = 'block';
                document.querySelector('.card-footer').style.display = 'none';
            } else {
                urgentNoData.style.display = 'none';
                document.querySelector('.card-footer').style.display = 'block';

                urgentOrdersList.forEach(order => {
                    const row = document.createElement('tr');
                    row.className = 'urgent-order';
                    row.innerHTML = `
                        <td>${order['订单ID'] || '-'}</td>
                        <td>${order['客户姓名'] || '-'}</td>
                        <td>${order['商品名称'] || '-'}</td>
                        <td>${order['客户国家'] || '-'}</td>
                        <td>${order['订单日期'] || '-'}</td>
                        <td><span class="badge bg-danger">${order.daysPending}天</span></td>
                        <td><span class="status-badge" style="background-color: ${statusColors[order['订单状态']] || '#6c757d'}; color: ${['待支付','待发货'].includes(order['订单状态']) ? '#fff' : '#000'}">
                            ${order['订单状态'] || '未知状态'}
                        </span></td>
                    `;
                    urgentTableBody.appendChild(row);
                });
            }

            // 更新图表
            updateCharts(orders);
        }

        // 更新图表数据
        function updateCharts(orders) {
            // 订单状态分布
            const statusCounts = {};
            orders.forEach(order => {
                const status = order['订单状态'] || '未知';
                statusCounts[status] = (statusCounts[status] || 0) + 1;
            });

            const statusLabels = Object.keys(statusCounts);
            const statusData = statusLabels.map(label => statusCounts[label]);
            const backgroundColors = statusLabels.map(label => statusColors[label] || '#6c757d');

            const statusCtx = document.getElementById('statusChart').getContext('2d');
            if (window.statusChartInstance) {
                window.statusChartInstance.destroy();
            }
            window.statusChartInstance = new Chart(statusCtx, {
                type: 'doughnut',
                data: {
                    labels: statusLabels,
                    datasets: [{
                        data: statusData,
                        backgroundColor: backgroundColors,
                        borderWidth: 0
                    }]
                },
                options: {
                    responsive: true,
                    plugins: {
                        legend: { position: 'right' },
                        title: { display: true, text: '订单状态分布' }
                    }
                }
            });

            // 国家销售分布
            const countryCounts = {};
            orders.forEach(order => {
                const country = order['客户国家'] || '未知';
                countryCounts[country] = (countryCounts[country] || 0) + 1;
            });

            // 取Top 10国家
            const sortedCountries = Object.entries(countryCounts)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 10);

            const countryLabels = sortedCountries.map(item => item[0]);
            const countryData = sortedCountries.map(item => item[1]);

            const countryCtx = document.getElementById('countryChart').getContext('2d');
            if (window.countryChartInstance) {
                window.countryChartInstance.destroy();
            }
            window.countryChartInstance = new Chart(countryCtx, {
                type: 'bar',
                data: {
                    labels: countryLabels,
                    datasets: [{
                        label: '订单数量',
                        data: countryData,
                        backgroundColor: '#4361ee',
                        borderWidth: 0
                    }]
                },
                options: {
                    responsive: true,
                    plugins: { legend: { display: false } },
                    scales: { y: { beginAtZero: true } }
                }
            });
        }