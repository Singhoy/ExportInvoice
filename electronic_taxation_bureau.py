#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@Author: Singhoy
@Created on: 2023/10/9
@file: electronic_taxation_bureau.py
@link: https://github.com/Singhoy/ExportInvoice
@Description:
电子税务局操作

输入账号密码登陆->我要办税-->开票业务->纸质发票业务->查看开具->开票日期选前一天和当天->查询->勾选点导出全量发票查询导出结果表
"""
__version__ = "0.0.1"

import asyncio
from datetime import timedelta
from os import path
from re import compile as rc
from time import time

from playwright.async_api import async_playwright, TimeoutError

HOME = "https://etax.guangdong.chinatax.gov.cn/xxmh/"
URI = "https://etax.guangdong.chinatax.gov.cn/xxmh/html/index_origin.html?gopage=true&m1=dzfpkj&m2=&fromWhere=&qxkzsx=&tabTitle=null&"
TAX_PAGE = "https://dppt.guangdong.chinatax.gov.cn:8443/invoice-query/invoice-query"


class TaxationBureau(object):
    def __init__(self, _log: object, cp: str, pz: dict):
        self.logger = _log
        self.cache_path = cp
        self.pz = pz
        self.wait_time = int(pz.get("查询等待时长", 120))
        self.ap = async_playwright()
        self.browser = None
        self.context = None
        self.page = None
        self.lu = rc(r'"(.*)"')
    
    async def __aenter__(self):
        ap = await self.ap.start()
        self.browser = await ap.chromium.launch(downloads_path=self.cache_path)
        self.context = await self.new_context()
        self.page = await self.new_page()
        await self.page.add_init_script('Object.defineProperties(navigator, {webdriver: {get: () => false}});')
        return self
    
    async def __aexit__(self, *args):
        await self._close()
    
    async def start(self):
        return await self.__aenter__()
    
    async def new_context(self):
        context = await self.browser.new_context(
            bypass_csp=True,
            permissions=["geolocation"],
            java_script_enabled=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
            offline=False,
            record_har_path="output.har"
        )
        await context.clear_cookies()
        return context
    
    async def new_page(self):
        return await self.context.new_page()
    
    async def _close(self):
        await self.browser.close()
        await self.ap.__aexit__()
    
    async def _run(self) -> str:
        res = await self.login()
        if res == 1:
            return
        return await self.export_data()
    
    async def export_data(self):
        """导出发票数据"""
        fail = 1
        for i in range(3):
            ok = await self.open_page()
            if ok != 1:
                fail = 0
                break
            self.logger.info("被踢下线了，5秒后重新来过...")
            await asyncio.sleep(5)
        if fail:
            self.logger.error("导出发票失败！")
            return
        self.logger.info("成功打开发票查询界面...")
        ks = self.pz["查询开始时间"] + timedelta(days=-1)
        js = self.pz["查询结束时间"] + timedelta(days=1)
        await self.change_date("开票日期起", ks)
        await self.change_date("开票日期止", js)
        await self.select_it()
        await self.page.screenshot(path=path.join(self.cache_path, f'查询条件{int(time() * 1000)}.png'))  # 保留查询条件截图
        async with self.page.expect_download(timeout=self.wait_time * 1000) as di:
            await self.page.get_by_text("导出全部").click()
        do = await di.value
        self.logger.debug(do)
        save_file = path.join(self.cache_path, "发票.xlsx")
        await do.save_as(save_file)
        return save_file
    
    async def select_it(self):
        """点击查询"""
        # 查询按钮
        a = await self.page.query_selector("button[type='submit']")
        await a.click()
        await self.page.wait_for_load_state()
        await self.page.wait_for_selector(".t-col.t-col-10")
        for i in range(5):
            button = await self.page.query_selector(
                '//div[@class="t-col t-col-10"]//button[contains(@class, "-primary")]')
            if button:
                await button.click()
                await self.page.wait_for_selector(".t-popup__content.t-dropdown", state='attached')
                break
            self.logger.info("等待数据加载...")
            await asyncio.sleep(3)
    
    async def change_date(self, name, nyr):
        """修改开票日期"""
        nyr = nyr.strftime("%Y-%m-%d")
        y, m, d = nyr.split("-")
        await self.page.get_by_placeholder(name).click()
        await self.page.wait_for_selector(".t-popup.t-date-picker__panel-container", state='attached')
        divs = await self.page.query_selector_all(".t-popup.t-date-picker__panel-container")
        for div in divs:
            style = await div.get_attribute("style")
            if style == "display: none;":
                continue
            table = await div.query_selector(".t-date-picker__panel-date")
            head = await table.query_selector(".t-date-picker__header")
            await self.exchange_ym(head, "month", f"{int(m)}月", "li")
            await self.exchange_ym(head, "year", y, "li > span > span")
            # t-date-picker__table
            body = await table.query_selector(".t-date-picker__table")
            await self.exchange_d(body, d)
    
    @staticmethod
    async def exchange_d(body, value):
        """修改日"""
        if value.startswith("0"):
            value = value[1:]
        # td class=t-date-picker__cell
        tds = await body.query_selector_all(
            "//td[contains(@class, 't-date-picker__cell') and not(contains(@class, '--additional'))]/div")
        for t in tds:
            text = await t.text_content()
            if text == value:
                await t.click()
                break
    
    @staticmethod
    async def exchange_ym(head, key, value, sl):
        """修改年、月"""
        _month = await head.query_selector(f".t-select__wrap.t-date-picker__header-controller-{key}")
        m_inp = await _month.query_selector("input")
        await m_inp.click()
        await _month.wait_for_selector(".t-select__list", state='attached')
        ul = await _month.query_selector("ul")
        lis = await ul.query_selector_all(sl)
        for i in lis:
            text = await i.text_content()
            if text.replace(" ", "") == value:
                await i.click()
                break
    
    async def open_page(self):
        """打开纸质发票业务页面"""
        url = None
        for i in range(3):
            url = await self.get_url2()
            if not url is None:
                break
            await asyncio.sleep(2)
        if url is None:
            self.logger.error("获取发票业务链接失败！")
            return 1
        await self.page.goto(url)
        self.logger.info(f"访问{url}")
        await self.page.wait_for_load_state("networkidle")
        await self.page.goto(TAX_PAGE)
        await self.page.wait_for_load_state("networkidle")
        page_url = self.page.url
        if "redirect_uri" in page_url:
            self.logger.warning(page_url)
            return 1
    
    async def get_url2(self):
        """获取纸质发票业务链接"""
        url = None
        for i in range(3):
            url = await self.get_url()
            if not url is None:
                break
            await asyncio.sleep(2)
        if url is None:
            self.logger.error("获取开票业务链接失败！")
            return
        await self.page.goto(url)
        self.logger.info(f"访问{url}")
        await self.page.wait_for_load_state("networkidle")
        # page.get_by_role("link", name="纸质发票业务")
        lis = await self.page.query_selector_all(".active")
        for i in lis:
            title = await i.inner_text()
            if title == "纸质发票业务":
                return await i.get_attribute("href")
    
    async def get_url(self):
        """获取开票业务链接"""
        # 我要办税
        # page.get_by_text("我要办税").click()
        await self.page.click("#wybs")
        # page.get_by_role("link", name="开票业务")
        lis = await self.page.query_selector_all("h4")
        for i in lis:
            title = await i.inner_text()
            if title != "开票业务":
                continue
            parent_element = await i.evaluate_handle("(element) => element.parentNode")  # 获取父节点
            onclick_value = await parent_element.evaluate(
                "(element) => element.getAttribute('onclick')")  # 获取 onclick 属性值
            if "开票业务" in onclick_value:
                a = self.lu.findall(onclick_value)
                b = a[0].split('"')
                c = b[0].rsplit("=", 1)
                return URI + c[0] + "=null&cdmc=" + b[-1]
    
    async def tc(self):
        try:
            # 等待弹窗出现
            await self.page.wait_for_selector('#layui-layer1', timeout=1000)
            # 关闭弹窗
            await self.page.evaluate('document.querySelector(".layui-layer-btn0").click()')
        except TimeoutError:
            pass
    
    async def login(self):
        """登录"""
        await self.page.goto(HOME)
        # 弹窗公告偶尔冒出来
        await self.tc()
        # page.get_by_role("link", name="登录", exact=True).click()
        await self.page.click(".loginico")
        await self.page.wait_for_load_state()
        await self.page.get_by_placeholder("统一社会信用代码/纳税人识别号").fill(self.pz["纳税人识别号"])
        await self.page.get_by_placeholder("居民身份证号码/手机号码/用户名").fill(self.pz["用户名"])
        await self.page.get_by_placeholder("个人用户密码").fill(self.pz["个人用户密码"])
        await self.verify()
        # page.get_by_role("button", name="登录").click()
        await self.page.click(".el-button.loginCls.el-button--primary")
        # 等页面跳转
        await asyncio.sleep(self.wait_time)
        await self.page.wait_for_load_state()
        # 获取整个页面的文本内容
        page_content = await self.page.evaluate('document.body.textContent')
        if "忘记密码" in page_content:
            self.logger.error("电子税局登录失败，请检查用户名和密码...")
            await self.page.screenshot(path=path.join(self.cache_path, f'登录失败{int(time() * 1000)}.png'))  # 保留失败截图
            return 1
        self.logger.info("电子税局登录成功...")
    
    async def verify(self):
        """滑块验证"""
        source_element = await self.page.query_selector(".handler.animate")
        # 获取源元素的位置和大小信息
        source_box = await source_element.bounding_box()
        source_x = source_box['x']
        source_y = source_box['y']
        source_width = source_box['width']
        source_height = source_box['height']
        # 计算目标位置
        target_x = source_x + 365
        target_y = source_y
        # 模拟拖拽操作
        await self.page.mouse.move(source_x + source_width / 2, source_y + source_height / 2)
        await self.page.mouse.down()
        while 1:
            await self.page.mouse.move(target_x + source_width / 2, target_y + source_height / 2)
            target_element = await self.page.query_selector(".handler.handler_ok_bg")
            if not target_element is None:
                break
            target_x += 15
        await self.page.mouse.up()
