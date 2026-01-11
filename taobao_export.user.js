// ==UserScript==
// @name         淘宝订单批量导出
// @namespace    http://tampermonkey.net/
// @version      1.1
// @description  在淘宝“已买到的宝贝”页面添加批量导出功能，支持自动翻页获取所有订单数据并导出为Excel。
// @author       AI Commander
// @match        https://buyertrade.taobao.com/trade/itemlist/list_bought_items.htm*
// @grant        GM_registerMenuCommand
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_xmlhttpRequest
// @require      https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js
// ==/UserScript==

(function() {
    'use strict';

    // 配置项
    const CONFIG = {
        nextPageDelay: 5000, // 翻页后等待时间(ms)，防止反爬和等待DOM渲染 (增加到5秒)
        scrollDelay: 1000,    // 滚动页面的间隔(ms) (增加到1秒)
        detailFetchDelay: 2000 + Math.random() * 3000, // 详情页抓取随机延迟 (2-5秒)
        concurrency: 1,      // 详情页抓取并发数 (降为1，最安全)
        batchSize: 5,        // 每抓取多少个订单暂停一次
        batchPause: 10000,   // 批次暂停时间 (ms)
        selectors: {
            orderContainer: "div[id^='shopOrderContainer_']",
            orderTime: "span[class*='shopInfoOrderTime']",
            orderId: "span[class*='shopInfoOrderId']",
            shopName: "a[class*='shopInfoName']",
            orderStatus: "span[class*='shopInfoStatus']",
            actualFee: "div[class*='priceReal--']", // 实付款区域
            itemInfo: "div[class*='itemInfo--']", // 商品信息行
            itemTitle: "a[class*='title--'] span[class*='titleText--']",
            itemSku: "div[class*='infoContent--']",
            itemPrice: "div[class*='itemInfoColPrice--']",
            itemQuantity: "div[class*='quantity--']",
            itemImage: "a[class*='image--']", // 商品主图链接
            nextPageBtn: "li.ant-pagination-next:not(.ant-pagination-disabled) button",
            detailLink: "a[class*='shopInfoOrderDetail--']", // 订单详情链接
            shippingFee: "div[class*='trade-price-container-block']" // 包含运费的容器
        },
        detailSelectors: {
            logistics: "div[class*='logisticsAddress--']", // 收货信息容器
            addressItem: "div[class*='addressItem--']",
            userInfo: "div[class*='userInfo--']", // 联系人信息
            infoContent: "div[class*='detailInfoContent--']", // 详情信息行
            infoTitle: "div[class*='detailInfoTitle--']", // 详情标题
            infoItem: "a[class*='detailInfoItem--']" // 详情内容
        }
    };

    // 存储所有订单数据
    let allOrdersData = [];
    let isExporting = false;
    let shouldStop = false;

    // 初始化界面
    function initUI() {
        const container = document.createElement('div');
        container.id = 'batch-export-container';
        container.style.cssText = `
            position: fixed;
            top: 100px;
            right: 20px;
            z-index: 99999;
            background: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            text-align: center;
            width: 200px;
        `;

        const btn = document.createElement('button');
        btn.id = 'batch-export-btn';
        btn.innerText = '批量导出所有订单';
        btn.style.cssText = `
            padding: 10px 20px;
            background: #ff4400;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            font-weight: bold;
            margin-bottom: 10px;
            width: 100%;
        `;
        btn.onclick = startBatchExport;

        const cancelBtn = document.createElement('button');
        cancelBtn.id = 'batch-cancel-btn';
        cancelBtn.innerText = '取消导出';
        cancelBtn.style.cssText = `
            padding: 10px 20px;
            background: #999;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            font-weight: bold;
            margin-bottom: 10px;
            width: 100%;
            display: none;
        `;
        cancelBtn.onclick = function() {
            if (confirm('确定要停止导出吗？已获取的数据将会被保存。')) {
                shouldStop = true;
                updateStatus('正在停止...');
            }
        };

        const status = document.createElement('div');
        status.id = 'export-status';
        status.style.cssText = `
            font-size: 12px;
            color: #666;
            margin-top: 10px;
            display: none;
            text-align: left;
        `;

        const progressBar = document.createElement('div');
        progressBar.id = 'export-progress';
        progressBar.style.cssText = `
            width: 100%;
            height: 6px;
            background: #f0f0f0;
            border-radius: 3px;
            margin-top: 10px;
            overflow: hidden;
            display: none;
        `;

        const progressFill = document.createElement('div');
        progressFill.id = 'progress-fill';
        progressFill.style.cssText = `
            height: 100%;
            background: #ff4400;
            width: 0%;
            transition: width 0.3s;
        `;

        progressBar.appendChild(progressFill);
        container.appendChild(btn);
        container.appendChild(cancelBtn);
        container.appendChild(status);
        container.appendChild(progressBar);
        document.body.appendChild(container);
    }

    function updateStatus(text, progress = null) {
        const statusEl = document.getElementById('export-status');
        const progressEl = document.getElementById('export-progress');
        const fillEl = document.getElementById('progress-fill');

        if (statusEl) {
            statusEl.style.display = 'block';
            statusEl.innerText = text;
        }

        if (progress !== null && progressEl && fillEl) {
            progressEl.style.display = 'block';
            fillEl.style.width = `${progress}%`;
        }
    }

    // 辅助函数：提取价格文本
    function extractPrice(container) {
        if (!container) return "0.00";
        
        let text = container.innerText.replace(/\s/g, '').replace('￥', '');
        // 去除常见前缀
        text = text.replace('实付款', '').replace('含运费', '').replace(':', '').replace('：', '');
        return text;
    }

    // 辅助函数：提取图片链接
    function extractImage(element) {
        if (!element) return "";
        let style = element.getAttribute('style');
        if (style) {
            let match = style.match(/url\("?(.*?)"?\)/);
            if (match && match[1]) {
                let url = match[1];
                if (url.startsWith('//')) {
                    url = 'https:' + url;
                }
                return url;
            }
        }
        if (element.tagName === 'IMG') {
             let url = element.src;
             if (url.startsWith('//')) {
                url = 'https:' + url;
            }
            return url;
        }
        return "";
    }

    // 抓取详情页数据
    function fetchOrderDetail(url) {
        if (!url) return Promise.resolve({});
        if (url.startsWith('//')) url = 'https:' + url;
        
        // console.log(`[淘宝导出] 开始抓取详情页: ${url}`);
        return Promise.resolve({}); // 暂时禁用详情页抓取，直接返回空对象

        // return new Promise((resolve) => {
        //     GM_xmlhttpRequest({
        //         method: "GET",
        //         url: url,
        //         onload: function(response) {
        //             try {
        //                 const parser = new DOMParser();
        //                 const doc = parser.parseFromString(response.responseText, 'text/html');

        //                 const details = {
        //                     '收件人信息': '',
        //                     '交易快照': '',
        //                     '支付宝交易号': '',
        //                     '创建时间': '',
        //                     '付款时间': '',
        //                     '发货时间': '',
        //                     '成交时间': ''
        //                 };

        //                 // 1) 收件人信息（物流地址块）
        //                 const logisticsEl = doc.querySelector(CONFIG.detailSelectors.logistics);
        //                 if (logisticsEl) {
        //                     const addressItem = logisticsEl.querySelector(CONFIG.detailSelectors.addressItem) || logisticsEl;
        //                     // 尝试更精确的提取：结构通常是 icon + div(div(addr) + div(user))
        //                     // addressItem 直接 innerText 也可以，但可能包含多余空白
        //                     let text = addressItem.innerText || '';
        //                     details['收件人信息'] = text.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
        //                 } else {
        //                     // 备选：尝试从“订单信息”列表里找“收货信息”
        //                     // 有时候物流块不显示，但下方的列表中有
        //                 }

        //                 // 2) 订单信息块（标题 + 值，右侧值通常是 .detailInfoItem--*** 的 a/div）
        //                 const rows = doc.querySelectorAll(CONFIG.detailSelectors.infoContent);
        //                 rows.forEach((row) => {
        //                     const titleEl = row.querySelector(CONFIG.detailSelectors.infoTitle);
        //                     if (!titleEl) return;
        //                     const title = titleEl.innerText.trim();

        //                     // 查找同级的值元素。注意：titleEl 本身也在一个 item 里，需要排除它。
        //                     // 结构通常是: div(item > title) + a(item > value)
        //                     const allItems = Array.from(row.querySelectorAll("[class*='detailInfoItem--']"));
        //                     // 找到不包含 titleEl 的那个 item
        //                     const valueEl = allItems.find(item => !item.contains(titleEl));

        //                     const getValueText = () => {
        //                         if (!valueEl) return '';
        //                         // 文本值
        //                         const txt = (valueEl.innerText || '').trim();
        //                         // 某些值是链接，优先用 href
        //                         if (title === '交易快照') {
        //                             if (valueEl.tagName === 'A' && valueEl.href) return valueEl.href;
        //                             const link = valueEl.querySelector('a');
        //                             if (link && link.href) return link.href;
        //                         }
        //                         return txt;
        //                     };

        //                     switch (title) {
        //                         case '收货信息': // 备选收货信息
        //                             if (!details['收件人信息']) {
        //                                 details['收件人信息'] = getValueText().replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
        //                             }
        //                             break;
        //                         case '交易快照':
        //                             details['交易快照'] = getValueText();
        //                             break;
        //                         case '支付宝交易号':
        //                             details['支付宝交易号'] = getValueText();
        //                             break;
        //                         case '创建时间':
        //                             details['创建时间'] = getValueText();
        //                             break;
        //                         case '付款时间':
        //                             details['付款时间'] = getValueText();
        //                             break;
        //                         case '发货时间':
        //                             details['发货时间'] = getValueText();
        //                             break;
        //                         case '成交时间':
        //                             details['成交时间'] = getValueText();
        //                             break;
        //                         default:
        //                             break;
        //                     }
        //                 });
                        
        //                 // console.log(`[淘宝导出] 解析详情完成 ${url}`, details);
        //                 resolve(details);
        //             } catch (e) {
        //                 // console.error('解析详情页失败:', url, e);
        //                 resolve({});
        //             }
        //         },
        //         onerror: function(err) {
        //             // console.error('抓取详情页网络错误:', url, err);
        //             resolve({});
        //         },
        //         ontimeout: function(err) {
        //             // console.error('抓取详情页超时:', url, err);
        //             resolve({});
        //         }
        //     });
        // });
    }

    // 解析当前页面订单
    function parseCurrentPage() {
        const orderContainers = document.querySelectorAll(CONFIG.selectors.orderContainer);
        // console.log(`当前页找到 ${orderContainers.length} 个订单`);
        
        const pageOrders = [];

        orderContainers.forEach(container => {
            try {
                // 订单基础信息
                const timeEl = container.querySelector(CONFIG.selectors.orderTime);
                const idEl = container.querySelector(CONFIG.selectors.orderId);
                const shopEl = container.querySelector(CONFIG.selectors.shopName);
                const statusEl = container.querySelector(CONFIG.selectors.orderStatus);
                const actualFeeEl = container.querySelector(CONFIG.selectors.actualFee);
                
                // 详情链接
                const detailLinkEl = container.querySelector(CONFIG.selectors.detailLink);
                const detailUrl = detailLinkEl ? detailLinkEl.href : '';

                const orderTime = timeEl ? timeEl.innerText.trim() : '';
                const orderId = idEl ? idEl.innerText.replace(/订单号[:：]/, '').trim() : '';
                const shopName = shopEl ? shopEl.innerText.trim() : '';
                const orderStatus = statusEl ? statusEl.innerText.trim() : '';
                const actualFee = extractPrice(actualFeeEl);

                // 运费提取
                // 运费通常在实付款列的下面，含有“含运费”字样
                let shippingFee = "0.00";
                const priceBlocks = container.querySelectorAll(CONFIG.selectors.shippingFee);
                priceBlocks.forEach(block => {
                    if (block.innerText.includes('含运费')) {
                        shippingFee = extractPrice(block);
                    }
                });

                // 商品列表 (一个订单可能有多个商品)
                const items = container.querySelectorAll(CONFIG.selectors.itemInfo);
                
                items.forEach(item => {
                    const titleEl = item.querySelector(CONFIG.selectors.itemTitle);
                    const skuEls = item.querySelectorAll(CONFIG.selectors.itemSku);
                    const priceEl = item.querySelector(CONFIG.selectors.itemPrice);
                    const qtyEl = item.querySelector(CONFIG.selectors.itemQuantity);
                    const imgEl = item.querySelector(CONFIG.selectors.itemImage);

                    const title = titleEl ? titleEl.innerText.trim() : '';
                    const sku = Array.from(skuEls).map(el => el.innerText.trim()).join('; ');
                    
                    // 单价和原价
                    // 如果存在 .trade-price-container-underline，则其为原价，另一个为单价（通常是实付单价）
                    // 这里的结构比较复杂，通常第一个显示的是当前价格，划线的是原价
                    // 但在脚本读取到的 html 中，priceEl 包含了所有价格块
                    // 简单处理：查找 priceEl 下所有的 priceWrap--... 
                    // 如果有多个，假设带有 underline 的是原价，不带的是单价
                    
                    let unitPrice = "0.00";
                    let originalPrice = "0.00";

                    if (priceEl) {
                        const prices = priceEl.querySelectorAll("div[class*='trade-price-container-block']");
                        prices.forEach(p => {
                            if (p.className.includes('underline')) {
                                originalPrice = extractPrice(p);
                            } else {
                                unitPrice = extractPrice(p);
                            }
                        });
                    }
                    
                    // 如果没有划线价，原价 = 单价
                    if (originalPrice === "0.00" || originalPrice === "") {
                        originalPrice = unitPrice;
                    }

                    const quantity = qtyEl ? qtyEl.innerText.replace('x', '').trim() : '1';
                    const mainImage = extractImage(imgEl);

                    pageOrders.push({
                        '订单号': orderId,
                        '下单时间': orderTime,
                        '店铺名称': shopName,
                        '商品标题': title,
                        '商品规格': sku,
                        '单价': unitPrice,
                        '原价': originalPrice,
                        '数量': quantity,
                        '实付款(含运费)': actualFee,
                        '运费': shippingFee,
                        '交易状态': orderStatus,
                        '商品主图': mainImage,
                        '详情链接': detailUrl, // 暂存，用于后续抓取
                        // 占位
                        '收件人信息': '',
                        '交易快照': '',
                        '支付宝交易号': '',
                        '创建时间': '',
                        '付款时间': '',
                        '发货时间': '',
                        '成交时间': ''
                    });
                });

            } catch (e) {
                // console.error('解析订单出错:', e, container);
            }
        });

        return pageOrders;
    }

    // 翻页逻辑
    async function goToNextPage() {
        const nextBtn = document.querySelector(CONFIG.selectors.nextPageBtn);
        if (nextBtn) {
            // 滚动到底部，确保按钮可见
            nextBtn.scrollIntoView();
            await new Promise(r => setTimeout(r, 500));
            
            nextBtn.click();
            return true;
        }
        return false;
    }

    // 延时函数
    const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

    // 并发控制器
    async function runConcurrent(tasks, limit) {
        const results = [];
        const executing = [];
        
        for (const task of tasks) {
            if (shouldStop) break;
            
            const p = Promise.resolve().then(() => task());
            results.push(p);
            
            if (limit <= tasks.length) {
                const e = p.then(() => executing.splice(executing.indexOf(e), 1));
                executing.push(e);
                if (executing.length >= limit) {
                    await Promise.race(executing);
                }
            }
        }
        return Promise.all(results);
    }

    // 开始批量导出
    async function startBatchExport() {
        if (isExporting) return;
        isExporting = true;
        shouldStop = false;
        allOrdersData = [];
        
        const btn = document.getElementById('batch-export-btn');
        const cancelBtn = document.getElementById('batch-cancel-btn');
        
        btn.style.display = 'none';
        cancelBtn.style.display = 'block';

        let hasNext = true;
        let pageCount = 1;

        try {
            while (hasNext) {
                if (shouldStop) break;

                updateStatus(`正在解析第 ${pageCount} 页...`);
                
                // 等待内容加载
                if (pageCount > 1) {
                    await delay(CONFIG.nextPageDelay);
                } else {
                    await delay(1000);
                }

                if (shouldStop) break;

                // 滚动页面
                window.scrollTo(0, document.body.scrollHeight);
                await delay(CONFIG.scrollDelay);
                window.scrollTo(0, 0);

                // 解析当前页基础数据
                const pageData = parseCurrentPage();
                
            // 抓取详情页数据
            if (pageData.length > 0) {
                updateStatus(`第 ${pageCount} 页：准备抓取订单详情 (${pageData.length} 个)...`);
                
                // 去重（同一订单号只需要抓取一次）
                const uniqueOrders = {}; // orderId -> url
                pageData.forEach(order => {
                    if (order['详情链接'] && !uniqueOrders[order['订单号']]) {
                        uniqueOrders[order['订单号']] = order['详情链接'];
                    }
                });

                const tasks = Object.entries(uniqueOrders).map(([oid, url], index) => {
                    return async () => {
                        if (shouldStop) return null;
                        
                        // 批次暂停机制
                        if (index > 0 && index % CONFIG.batchSize === 0) {
                            updateStatus(`已抓取 ${index} 个，暂停 ${CONFIG.batchPause/1000} 秒防拦截...`);
                            await delay(CONFIG.batchPause);
                        }

                        updateStatus(`第 ${pageCount} 页：正在抓取订单 ${oid} 详情 (${index + 1}/${Object.keys(uniqueOrders).length})...`);
                        await delay(CONFIG.detailFetchDelay); 
                        const detail = await fetchOrderDetail(url);
                        return { oid, detail };
                    };
                });

                // 并发执行任务
                const results = await runConcurrent(tasks, CONFIG.concurrency);
                
                // 建立结果映射
                const detailMap = {};
                results.forEach(res => {
                    if (res) {
                        detailMap[res.oid] = res.detail;
                    }
                });

                // 回填数据
                pageData.forEach(order => {
                    const oid = order['订单号'];
                    if (detailMap[oid]) {
                        Object.assign(order, detailMap[oid]);
                    }
                });
                
                if (shouldStop) {
                    // 如果停止，也要保留已处理的部分数据
                    allOrdersData = allOrdersData.concat(pageData);
                    break;
                }

                allOrdersData = allOrdersData.concat(pageData);
                updateStatus(`已收集 ${allOrdersData.length} 条订单记录`, 50);
            }

                if (shouldStop) break;

                // 尝试翻页
                hasNext = await goToNextPage();
                if (hasNext) {
                    pageCount++;
                } else {
                    updateStatus('已到达最后一页，准备生成文件...');
                }
            }
        } catch (e) {
            // console.error('导出过程出错:', e);
            updateStatus('导出出错，尝试保存已获取数据...');
        }

        // 导出文件
        if (allOrdersData.length > 0) {
            exportExcel(allOrdersData);
            if (shouldStop) {
                updateStatus(`已取消导出。共导出 ${allOrdersData.length} 个订单`, 100);
            } else {
                updateStatus(`导出完成！共导出 ${allOrdersData.length} 个订单`, 100);
            }
        } else {
            updateStatus('未获取到任何订单数据');
        }
        
        isExporting = false;
        shouldStop = false;
        btn.style.display = 'block';
        cancelBtn.style.display = 'none';
        btn.disabled = false;
        btn.innerText = '批量导出所有订单';
    }

    // 导出 Excel
    function exportExcel(data) {
        if (data.length === 0) {
            alert('没有找到订单数据');
            return;
        }

        // 移除不必要的列（如详情链接）
        const exportData = data.map(item => {
            const { '详情链接': _, ...rest } = item;
            return rest;
        });

        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "订单数据");

        // 设置列宽
        const colWidths = [
            { wch: 25 }, // 订单号
            { wch: 20 }, // 下单时间
            { wch: 20 }, // 店铺名称
            { wch: 50 }, // 商品标题
            { wch: 30 }, // 商品规格
            { wch: 10 }, // 单价
            { wch: 10 }, // 原价
            { wch: 8 },  // 数量
            { wch: 15 }, // 实付款
            { wch: 10 }, // 运费
            { wch: 15 }, // 交易状态
            { wch: 50 }, // 商品主图
            { wch: 50 }, // 收件人信息
            { wch: 50 }, // 交易快照
            { wch: 30 }, // 支付宝交易号
            { wch: 20 }, // 创建时间
            { wch: 20 }, // 付款时间
            { wch: 20 }, // 发货时间
            { wch: 20 }  // 成交时间
        ];
        ws['!cols'] = colWidths;

        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
        XLSX.writeFile(wb, `淘宝订单导出_${timestamp}.xlsx`);
    }

    // 启动
    window.addEventListener('load', function() {
        setTimeout(initUI, 2000); // 延迟初始化，确保页面基本加载完成
    });

})();
