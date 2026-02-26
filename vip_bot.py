import asyncio
from playwright.async_api import async_playwright
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
import os
import random
import json

get_page_number_js = r"""
() => {
  // 查找包含总页数的元素
  const span = document.querySelector('span.total-item-nums');
  if (!span) return null;
  // 提取出"共27页"这种文本里的数字部分
  const match = span.textContent.match(/共(\d+)页/);
  return match ? parseInt(match[1], 10) : null;
}
"""


async def get_page_number(keyword, page):
    """
    获取某个商品的页数，复用已有的 page
    """
    await page.goto(f"https://category.vip.com/suggest.php?keyword={keyword}")
    print(f"Current URL: {page.url}")
    page_number = await page.evaluate(get_page_number_js)
    print(f"共：{page_number} 页")
    return page_number


get_items_of_page_js = r"""
() => {
  // 获取所有商品元素
  const items = Array.from(document.querySelectorAll('div.c-goods-item.J-goods-item.c-goods-item--auto-width[data-product-id]'));
  return items.map(item => {
    // 商品ID
    const productId = item.getAttribute('data-product-id');

    // 商品链接
    const aTag = item.querySelector('a[href*="detail"]');
    // 有部分链接是 // 开头，加上协议
    let href = aTag ? aTag.getAttribute('href') : null;
    if (href && href.startsWith("//")) {
      href = "https:" + href;
    }

    // 商品底部信息块
    const bottom = item.querySelector('div.c-goods-item-bottom');
    if (!bottom) return null;

    // 主要价格（特卖价）
    const salePriceDiv = bottom.querySelector('div.c-goods-item__sale-price');
    let salePrice = null;
    if (salePriceDiv) {
      const priceText = salePriceDiv.textContent.replace(/[^\d.]/g, '');
      salePrice = priceText ? parseFloat(priceText) : null;
    }

    // 原价（市场价）
    const marketPriceDiv = bottom.querySelector('div.c-goods-item__market-price');
    let marketPrice = null;
    if (marketPriceDiv) {
      const priceText = marketPriceDiv.textContent.replace(/[^\d.]/g, '');
      marketPrice = priceText ? parseFloat(priceText) : null;
    }

    // 折扣
    const discountDiv = bottom.querySelector('div.c-goods-item__discount');
    let discount = null;
    if (discountDiv) {
      discount = discountDiv.textContent.trim();
    }

    // 商品名称
    const nameDiv = bottom.querySelector('div.c-goods-item__name');
    let name = null;
    if (nameDiv) {
      name = nameDiv.textContent.trim();
    }

    return {
      productId,
      href,
      salePrice,
      marketPrice,
      discount,
      name
    };
  }).filter(x => x !== null);
}
"""

get_detail_info_js = r"""
() => {
  const result = {
    sizes: [],
    productCode: ''
  };

  // 抓取所有尺码
  const sizeElements = document.querySelectorAll('span.size-list-item-name');
  if (sizeElements.length > 0) {
    result.sizes = Array.from(sizeElements).map(el => el.textContent.trim());
  }

  // 抓取商品编码 - 方法1：在tbody中查找
  const tbody = document.querySelector('tbody.J_dc_table');
  if (tbody) {
    const thElements = tbody.querySelectorAll('th.dc-table-tit');
    for (let th of thElements) {
      if (th.textContent.includes('商品编码')) {
        // 找到最近的td兄弟元素
        let nextEl = th.nextElementSibling;
        while (nextEl && nextEl.tagName !== 'TD') {
          nextEl = nextEl.nextElementSibling;
        }
        if (nextEl) {
          result.productCode = nextEl.textContent.trim();
          break;
        }
      }
    }
  }

  // 如果还没有找到，尝试方法2：查找整个页面中包含"商品编码"的元素
  if (!result.productCode) {
    const allText = document.body.innerText;
    const match = allText.match(/商品编码[：:]\s*([A-Z0-9]+)/);
    if (match) {
      result.productCode = match[1];
    }
  }

  return result;
}
"""

def get_progress_file(keyword):
    """
    获取进度文件路径
    """
    return f"data/{keyword}_progress.json"


def load_progress(keyword):
    """
    读取抓取进度
    返回：已完成的最后一页页码，如果没有进度则返回0
    """
    progress_file = get_progress_file(keyword)
    if os.path.exists(progress_file):
        try:
            with open(progress_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                last_page = data.get('last_completed_page', 0)
                print(f"📖 读取进度：上次已完成第 {last_page} 页")
                return last_page
        except Exception as e:
            print(f"⚠️  读取进度文件失败: {e}")
    return 0


def save_progress(keyword, page):
    """
    保存抓取进度
    """
    progress_file = get_progress_file(keyword)

    # 确保目录存在
    os.makedirs(os.path.dirname(progress_file), exist_ok=True)

    data = {
        'keyword': keyword,
        'last_completed_page': page,
        'last_update_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    try:
        with open(progress_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"💾 进度已保存：第 {page} 页")
    except Exception as e:
        print(f"⚠️  保存进度失败: {e}")


def write_items_to_excel(items, keyword, page, output_file='products.xlsx'):
    """
    将items数据写入Excel文件，每个尺码对应一行数据
    每页创建独立的Excel文件
    """
    # 确保输出目录存在
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 创建新文件（每页一个独立文件）
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "商品数据"

    # 设置表头
    headers = ["标题", "市场价", "尺码", "商品编码", "折扣", "品牌", "产品id", "详情页"]
    ws.append(headers)

    # 设置表头样式
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    # 写入数据 - 每个尺码对应一行
    total_rows = 0
    for item in items:
        # 获取尺码列表和商品编码
        sizes = item.get('sizes', [])
        product_code = item.get('productCode', '')

        # 如果没有尺码，则创建一条记录（尺码为空）
        if not sizes:
            sizes = ['']

        # 为每个尺码创建一行
        for size in sizes:
            row = [
                item.get('name', ''),           # 标题
                item.get('marketPrice', ''),    # 市场价
                size,                           # 尺码
                product_code,                   # 商品编码
                item.get('discount', ''),       # 折扣
                keyword,                        # 品牌 (keyword)
                item.get('productId', ''),      # 产品id
                item.get('href', '')            # 详情页链接
            ]
            ws.append(row)
            total_rows += 1

    # 自动调整列宽
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

    # 保存文件
    wb.save(output_file)
    print(f"✓ Excel文件已保存: {output_file}")
    print(f"✓ 共导入 {total_rows} 条商品数据")


async def human_scroll(page):
    """
    模拟真实用户的滑动，每次滑动到页面最底部，然后等待3秒加载内容
    连续2次页面高度不变则停止滚动
    """
    last_height = 0
    no_change_count = 0

    while no_change_count < 2:  # 连续2次高度不变则停止滚动
        # 滑动到页面最底部
        await page.evaluate('window.scrollTo(0, document.documentElement.scrollHeight)')
        print(f"  已滑动到页面底部")

        # 等待内容加载
        await asyncio.sleep(1.5)

        # 获取当前页面高度
        current_height = await page.evaluate('document.documentElement.scrollHeight')

        if current_height == last_height:
            no_change_count += 1
            print(f"  页面高度未变化 ({no_change_count}/2) - 高度: {current_height}px")
        else:
            no_change_count = 0
            print(f"  页面高度已变化 - 上次: {last_height}px, 当前: {current_height}px")

        last_height = current_height

    print("✓ 页面滚动完成，无更多内容加载")


async def check_captcha(page_obj):
    """
    检测页面是否出现唯品会安全校验弹窗，如果有则暂停等待用户手动完成验证
    通过检测 div.vipsc_modal.vipsc_v_show 容器是否可见来判断
    """
    try:
        element = await page_obj.query_selector('div.vipsc_modal.vipsc_v_show')
        if element and await element.is_visible():
            print(f"\n⚠️  检测到安全校验弹窗，请在浏览器中完成验证")
            print(f"👉 完成验证后，回到终端按回车键继续...")
            await asyncio.get_event_loop().run_in_executor(None, input)
            print(f"✓ 继续抓取")
            return True
    except:
        pass

    return False


async def get_detail_info(page_obj, item, max_retries=3):
    """
    打开详情页并获取尺码和商品编码信息
    带重试机制，遇到验证码会等待用户处理
    """
    if not item.get('href'):
        return item

    for attempt in range(max_retries):
        try:
            await page_obj.goto(item['href'], wait_until='load', timeout=10000)

            # 检测验证码
            has_captcha = await check_captcha(page_obj)
            if has_captcha:
                # 验证码完成后重新加载页面
                await page_obj.reload(wait_until='load')

            detail_info = await page_obj.evaluate(get_detail_info_js)
            item['sizes'] = detail_info.get('sizes', [])
            item['productCode'] = detail_info.get('productCode', '')
            print(f"✓ 获取详情: {item['name'][:20]} - 尺码数: {len(item['sizes'])} - 编码: {item['productCode']} - 链接：{item['href']}")
            return item

        except Exception as e:
            error_msg = str(e)
            print(f"✗ 获取详情失败 (第{attempt + 1}次): {item['name'][:20]} - {error_msg}")

            # 如果是超时错误，增加延迟后重试
            if 'timeout' in error_msg.lower() or attempt < max_retries - 1:
                wait_time = (attempt + 1) * 5  # 递增延迟：5s, 10s, 15s
                print(f"  等待 {wait_time} 秒后重试...")
                await asyncio.sleep(wait_time)
            else:
                print(f"  放弃获取该商品详情")
                return item

    return item


async def setup_route(page):
    """
    为页面设置资源拦截，屏蔽图片、字体、媒体（captcha 相关的除外）
    """
    async def block_resources(route):
        if route.request.resource_type in ["image", "font", "media", "stylesheet"] and "captcha" not in route.request.url:
            await route.abort()
        else:
            await route.continue_()

    await page.route("**/*", block_resources)


async def get_items_of_page(keyword, page_num, page):
    """
    获取某个商品某页的数据，用同一个 tab 完成列表抓取和详情抓取
    """
    await page.goto(f"https://category.vip.com/suggest.php?keyword={keyword}&page={page_num}")
    print(f"Current URL: {page.url}")

    print("正在加载页面数据，请稍候...")
    await human_scroll(page)

    items = await page.evaluate(get_items_of_page_js)
    print(f"✓ 获取到 {len(items)} 个商品")

    if not items:
        print("⚠️ 本页没有获取到商品数据")
        return []

    print(f"开始获取商品详情，共 {len(items)} 个...")

    for i, item in enumerate(items):
        print(f"  [{i + 1}/{len(items)}] ", end="")
        await get_detail_info(page, item)

    today = datetime.today().strftime('%Y-%m-%d')
    output_file = f"data/{keyword}_{today}_page_{page_num}.xlsx"
    write_items_to_excel(items, keyword, page_num, output_file)

    save_progress(keyword, page_num)

    return items


async def main():
    keyword = "阿迪达斯"

    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp("http://localhost:9222")
        context = browser.contexts[0]
        page = context.pages[0]
        await setup_route(page)

        # 获取总页数（复用同一个 page）
        total_page = await get_page_number(keyword=keyword, page=page)

        if not total_page:
            print("⚠️ 未获取到总页数")
            return

        # 读取上次的进度
        last_completed_page = load_progress(keyword)
        start_page = last_completed_page + 1

        if start_page > total_page:
            print(f"\n✓ {keyword} 的所有数据已抓取完成！")
            return

        print(f"\n开始抓取 {keyword} 的商品数据")
        print(f"总页数: {total_page} | 起始页: {start_page} | 剩余页数: {total_page - start_page + 1}")
        print("=" * 60)

        for page_num in range(start_page, total_page + 1):
            print(f"\n【第 {page_num}/{total_page} 页】")
            try:
                await get_items_of_page(keyword=keyword, page_num=page_num, page=page)
            except Exception as e:
                print(f"✗ 第 {page_num} 页抓取失败: {e}")
                print(f"⚠️  已保存进度到第 {page_num - 1} 页，下次运行将从第 {page_num} 页继续")
                break

    print("\n" + "=" * 60)
    if last_completed_page > 0:
        print(f"✓ {keyword} 的数据抓取完成！（从第 {start_page} 页继续）")
    else:
        print(f"✓ {keyword} 的所有数据抓取完成！")


async def test_first_page():
    """
    测试阿迪达斯第一页的抓取情况
    """
    keyword = "阿迪达斯"
    print("=" * 60)
    print(f"测试开始：{keyword} 第1页抓取")
    print("=" * 60)

    start_time = datetime.now()

    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp("http://localhost:9222")
        context = browser.contexts[0]
        page = context.pages[0]
        await setup_route(page)

        items = await get_items_of_page(keyword=keyword, page_num=1, page=page)


    elapsed = (datetime.now() - start_time).total_seconds()

    # 统计结果
    total = len(items)
    with_sizes = sum(1 for item in items if item.get('sizes'))
    with_code = sum(1 for item in items if item.get('productCode'))
    total_sizes = sum(len(item.get('sizes', [])) for item in items)

    print("\n" + "=" * 60)
    print("测试结果统计")
    print("=" * 60)
    print(f"总耗时        : {elapsed:.1f} 秒")
    print(f"抓取商品数    : {total}")
    print(f"成功获取尺码  : {with_sizes} / {total}")
    print(f"成功获取编码  : {with_code} / {total}")
    print(f"总尺码条目数  : {total_sizes}")
    if total > 0:
        print(f"平均每件耗时  : {elapsed / total:.1f} 秒")
    print("=" * 60)


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "test":
        asyncio.run(test_first_page())
    else:
        asyncio.run(main())
