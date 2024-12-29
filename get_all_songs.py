from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from urllib.parse import unquote
import sqlite3, os, yaml, re

# 配置读取
config_file = 'config.yaml'
def read_config() -> dict:
    with open(config_file, 'r', encoding='utf-8') as file:
        return yaml.safe_load(file)

config = read_config()
db_name = config.get('db_name')
table_name_level = config.get('table_name_level')
table_name_new = config.get('table_name_new')
table_name_lun_new = config.get('table_name_lun_new')
pic_dir = config.get('pic_dir')
full_load = config.get('full_load')
specify_load = config.get('specify_load')
specify_urls = config.get('specify_urls')
retry_times = config.get('retry_times')
difficulty_names = config.get('difficulty_names')
wiki_songs_url = config.get('wiki_songs_url')
xpath_all_song_link = config.get('xpath_all_song_link')
xpath_new_song_link = config.get('xpath_new_song_link')
xpath_new_lun_song_link = config.get('xpath_new_lun_song_link')
xpath_song_title = config.get('xpath_song_title')
xpath_song_img = config.get('xpath_song_img')
xpath_song_difficulty_table = config.get('xpath_song_difficulty_table')
main_load_time = config.get('main_load_time')
song_load_time = config.get('song_load_time')

# 根据url获取song_id
def get_id_from_url(url: str) -> str:
    illegal_chars = r'[<>:"/\\|?*]'
    return re.sub(illegal_chars, '-', unquote(url.split('/')[-1]))

# SQL
sql_difficulty_cols = ['bas', 'adv', 'exp', 'mas', 'lun']
sql_drop_table = '''DROP TABLE IF EXISTS {table}'''
sql_create_table_level = f'''CREATE TABLE IF NOT EXISTS {table_name_level} (
                     id   TEXT    PRIMARY KEY     NOT NULL,
                     name    TEXT    NOT NULL,
                     {sql_difficulty_cols[0]}    REAL,
                     {sql_difficulty_cols[1]}    REAL,
                     {sql_difficulty_cols[2]}    REAL,
                     {sql_difficulty_cols[3]}    REAL,
                     {sql_difficulty_cols[4]}    REAL
                     )'''
sql_select_all_ids = f'''SELECT id FROM {table_name_level}'''
sql_insert_level = '''INSERT OR IGNORE INTO {table}(id, name) VALUES ("{id}", "{name}")'''
sql_update_level = '''UPDATE {table} SET {key}={value} WHERE id="{id}"'''
sql_create_table_new = '''CREATE TABLE {table} (
                     id   TEXT    PRIMARY KEY     NOT NULL
                     )'''
sql_insert_new = '''INSERT OR IGNORE INTO {table}(id) VALUES ("{id}")'''

# 定数推断
def infer_level(song_level: str, measured_level: str) -> float:
    try:
        return float(measured_level)
    except ValueError:
        pass
    # 通过歌曲等级推断最小定数
    flag_plus = False
    if song_level.endswith('+'):
        flag_plus = True
        song_level = song_level[:-1]
    try:
        infer_level = float(song_level)
        if flag_plus:
            infer_level += 0.7
        return infer_level
    except ValueError:
        return 0.0

if __name__ == '__main__':
    # 数据库初始化
    db_conn = sqlite3.connect(db_name)
    db_cursor = db_conn.cursor()
    for table in [table_name_new, table_name_lun_new]:
        db_cursor.execute(sql_drop_table.format(table=table))
        db_cursor.execute(sql_create_table_new.format(table=table))
    if full_load:
        db_cursor.execute(sql_drop_table.format(table=table_name_level))
    db_cursor.execute(sql_create_table_level)
    db_conn.commit()

    # 曲绘文件夹初始化
    try:
        os.mkdir(pic_dir)
    except FileExistsError:
        pass
        
    driver = webdriver.Chrome()
    driver.set_page_load_timeout(main_load_time)
    try:
        driver.get(wiki_songs_url)
    except TimeoutException:
        pass
    print('主页面加载完成')

    # 获取新歌铺面链接
    new_links = driver.find_elements(By.XPATH, xpath_new_song_link)
    new_ids = [get_id_from_url(link.get_attribute('href')) for link in new_links]
    # 去重
    new_ids = set(new_ids)
    for id in new_ids:
        db_cursor.execute(sql_insert_new.format(table=table_name_new, id=id))
    db_conn.commit()
    print(f'发现{len(new_ids)}首新版本常规曲')

    new_lun_links = driver.find_elements(By.XPATH, xpath_new_lun_song_link)
    new_lun_ids = [get_id_from_url(link.get_attribute('href')) for link in new_lun_links]
    # 去重
    new_lun_ids = set(new_lun_ids)
    for id in new_lun_ids:
        db_cursor.execute(sql_insert_new.format(table=table_name_lun_new, id=id))
    db_conn.commit()
    print(f'发现{len(new_lun_ids)}首新版本白谱')

    if not specify_load:
        # 获取全部铺面链接
        song_links = driver.find_elements(By.XPATH, xpath_all_song_link)
        song_urls = [link.get_attribute('href') for link in song_links]
        # 去重
        song_urls = list(set(song_urls))
    else:
        song_urls = specify_urls
    print(f'将加载{len(song_urls)}首歌曲')

    # 获取已有铺面信息
    db_cursor.execute(sql_select_all_ids)
    queried_ids = [row[0] for row in db_cursor.fetchall()]

    # 获取铺面信息
    driver.set_page_load_timeout(song_load_time)
    failed_urls = []
    for song_url in song_urls:
        song_id = get_id_from_url(song_url)
        if song_id in queried_ids:
            continue
        done = False
        for i in range(retry_times):
            try:
                try:
                    driver.get(song_url)
                except TimeoutException:
                    pass
                song_name = driver.find_element(By.XPATH, xpath_song_title).text
                print('正在读取：{name}'.format(name=song_name))
                # 获取铺面封面
                img = driver.find_element(By.XPATH, xpath_song_img)
                img_filename = f"./{pic_dir}/{song_id}.png"
                img.screenshot(img_filename)
                img_size = os.path.getsize(img_filename)
                if img_size < 5*1024:
                    raise Exception("图片为空")
                # 获取定数表
                db_cursor.execute(sql_insert_level.format(table=table_name_level, name=song_name, id=song_id))
                level_table = driver.find_element(By.XPATH, xpath_song_difficulty_table)
                for row_index in range(int(level_table.get_attribute('childElementCount'))):
                    current_row = level_table.find_element(By.XPATH, f'./tr[{row_index+1}]')
                    song_difficulty = current_row.find_element(By.XPATH, './th').text
                    try:
                        difficulty_index = difficulty_names.index(song_difficulty)
                    except Exception:
                        print(f'未知铺面难度{song_difficulty}')
                        continue
                    song_level = current_row.find_element(By.XPATH, './td[1]').text
                    measured_level = current_row.find_element(By.XPATH, './td[4]').text
                    level = infer_level(song_level, measured_level)
                    print(f'定数推断：{song_difficulty} | {song_level} | {measured_level} -> {level}')
                    if abs(level) <= 1e-6:
                        print('跳过0定数铺面')
                        continue
                    db_cursor.execute(sql_update_level.format(table=table_name_level, id=song_id, key=sql_difficulty_cols[difficulty_index], value=level))
                db_conn.commit()
                done = True
                break
            except Exception as e:
                print(e)
        if not done:
            print("加载失败")
            failed_urls.append(song_url)

    driver.quit()
    if len(failed_urls) > 0:
        print('以下铺面加载失败：')
        for id in failed_urls:
            print(f'- {id}')
        print('解决方式：关闭全量加载选项并重试')
        print('或将url列表复制到配置文件中的specify_urls，并将specify_load设置为true并重试')