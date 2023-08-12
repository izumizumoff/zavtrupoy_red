import sqlite3
import datetime
import pysftp
import config
from base64 import decodebytes
import paramiko

def del_all_table():
    # delete all tables
    conn = sqlite3.connect('zavtrupoy.db')
    cur = conn.cursor()
    cur.execute('drop table if exists acts')
    cur.execute('drop table if exists roles')
    cur.execute('drop table if exists actors')


def start():
    # create new default empty tables
    conn = sqlite3.connect('zavtrupoy.db')
    conn.execute('PRAGMA foreign_keys=ON')
    cur = conn.cursor()

    query_acts = 'CREATE TABLE acts (name TEXT NOT NULL UNIQUE PRIMARY KEY,space TEXT,assist TEXT,time INTEGER DEFAULT 90,date DATETIME);'
    query_actors = 'CREATE TABLE actors (name_actor TEXT NOT NULL UNIQUE PRIMARY KEY,gender TEXT,age DATETIME,start DATETIME,photo TEXT,bio TEXT, state TEXT, short_name TEXT);'
    # в таблице roles название спектакля зависит от назваия спектакля в acts 
    # (изменение и удаление)
    # в таблице roles имя актера зависит от имени актера в actors
    # (изменение и при удалении из actors - NULL)
    query_roles = 'CREATE TABLE roles (act_from_acts TEXT,role TEXT,actor TEXT,bal INT, short_role TEXT, FOREIGN KEY (act_from_acts) REFERENCES acts(name) ON DELETE CASCADE ON UPDATE CASCADE, FOREIGN KEY(actor) REFERENCES actors(name_actor) ON DELETE SET NULL ON UPDATE CASCADE);'

    cur.execute(query_acts)
    cur.execute(query_actors)
    cur.execute(query_roles)
    
    conn.commit()
    conn.close()


def add_acts(name, space, assist, time, date):
    # add data to acts table
    # name: str --> название спектакля
    # space: str --> пространство спектакля
    # assist: str --> помощник режиссера спектакля
    # time: int --> продолжительность спектакля в минутах
    # date: str (for datetime) --> дата премьеры
    conn = sqlite3.connect('zavtrupoy.db')
    conn.execute('PRAGMA foreign_keys=ON')
    cur = conn.cursor()
    data = (name, space, assist, time, date)
    query = 'insert into acts values (?, ?, ?, ?, ?)'
    cur.execute(query, data)
    conn.commit()
    conn.close()


def add_actors(name_actor, gender, age, start, photo, bio, state, short_name):
    # add data to actors table
    # name_actor: str --> имя актера
    # gender str: --> пол (м/ж)
    # age: str (for datetime) --> дата рождения актера
    # start: str (for datetime) --> дата начала работы в театре
    # photo: str --> ссылка на фото
    # bio: str --> краткая информация об актере
    # state: str --> state: постоянный актер в труппе; temp: приглашенный артист
    conn = sqlite3.connect('zavtrupoy.db')
    conn.execute('PRAGMA foreign_keys=ON')
    cur = conn.cursor()
    data = (name_actor, gender, age, start, photo, bio, state, short_name)
    query = 'insert into actors values (?, ?, ?, ?, ?, ?, ?, ?)'
    cur.execute(query, data)
    conn.commit()
    conn.close()

def add_roles(act_from_acts, role, actor, bal, short_role) -> None:
    # add data to roles table
    # act_from_acts: str --> название спектакля
    # role: str --> название роли
    # actor: str --> имя актера
    # bal: --> количество баллов за роль
    conn = sqlite3.connect('zavtrupoy.db')
    conn.execute('PRAGMA foreign_keys=ON')
    cur = conn.cursor()
    data = (act_from_acts, role, actor, bal, short_role)
    query = 'insert into roles values (?, ?, ?, ?, ?)'
    cur.execute(query, data)
    conn.commit()
    conn.close()

def update_act(column, new, reference):
    # update data of column in table
    # table: str --> назавание таблицы
    # column: str --> название колонны для изменения
    # row: str --> название колонны ориентира (какую строку отслеживать)
    # new: str --> новое значение
    # reference: str --> значение колонны ориентира
    conn = sqlite3.connect('zavtrupoy.db')
    conn.execute('PRAGMA foreign_keys=ON')
    cur = conn.cursor()
    data = (new, reference)
    query = f'update acts set {column}=? where name=?'
    cur.execute(query, data)
    conn.commit()
    conn.close()

def update_actor(column, new, reference):
    # update data of column in table
    # column: str --> название колонны для изменения
    # new: str --> новое значение
    # reference: str --> значение колонны ориентира
    conn = sqlite3.connect('zavtrupoy.db')
    conn.execute('PRAGMA foreign_keys=ON')
    cur = conn.cursor()
    data = (new, reference)
    query = f'update actors set {column}=? where name_actor=?'
    cur.execute(query, data)
    conn.commit()
    conn.close()

def update_role(column, new, act, role, actor):
    # update data of column in table
    # column: str --> название колонны для изменения
    # new: str --> новое значение
    # reference: str --> значение колонны ориентира
    conn = sqlite3.connect('zavtrupoy.db')
    conn.execute('PRAGMA foreign_keys=ON')
    cur = conn.cursor()
    data = (new, act, role, actor)
    query = f'update roles set {column}=? where act_from_acts=? and role=? and actor=?'
    cur.execute(query, data)
    conn.commit()
    conn.close()

def del_act(reference):
    # удалить строку из таблицы по названию спектаклю
    conn = sqlite3.connect('zavtrupoy.db')
    conn.execute('PRAGMA foreign_keys=ON')
    cur = conn.cursor()
    data = (reference,)
    query = f'DELETE from acts where name=?'
    cur.execute(query, data)
    conn.commit()
    conn.close()

def del_actor(reference):
    # удалить строку из таблицы по имени актера
    conn = sqlite3.connect('zavtrupoy.db')
    conn.execute('PRAGMA foreign_keys=ON')
    cur = conn.cursor()
    data = (reference,)
    query = f'DELETE from actors where name_actor=?'
    cur.execute(query, data)
    conn.commit()
    conn.close()

def del_role(act, role):
    # удалить строку из таблицы по роли и названию спектаклю
    conn = sqlite3.connect('zavtrupoy.db')
    conn.execute('PRAGMA foreign_keys=ON')
    cur = conn.cursor()
    data = (act, role)
    query = f'DELETE from roles where act_from_acts=? and short_role=?'
    cur.execute(query, data)
    conn.commit()
    conn.close()

def get_data(table, column='*', filter=None):
    # get data from tables
    # table: str --> название таблицы данные которой нужны
    # column [optional]: str --> название колонны данные которой нужны
    # filter [optional]: list --> [название колонны для референса, название референса]
    ## референс: все данные строк у которых в столбце референса такое значение референса
    conn = sqlite3.connect('zavtrupoy.db')
    conn.execute('PRAGMA foreign_keys=ON')
    cur = conn.cursor()
    if filter and len(filter) == 2:
        query = f'SELECT {column} FROM {table} WHERE {filter[0]}=?'
        data = (filter[1],)
        cur.execute(query, data)
    elif filter and len(filter) == 4:
        query = f'SELECT {column} FROM {table} WHERE {filter[0]}=? and {filter[2]}=?'
        data = (filter[1],filter[3])
        cur.execute(query, data)
    else:
        query = f'SELECT {column} FROM {table}'
        cur.execute(query)
    return list(cur.fetchall())

def date(d, time='00:00') -> datetime:
    # convert string "dddd-mm-dd" to datetime object
    # d: str --> строка даты в формате "2000-1-1" = 1 января 2000 года
    result = datetime.datetime(*tuple((int(x) for x in d.split("-"))) + tuple((int(x) for x in time.split(":"))))
    return result

def back_up():
    # отправляет по sftp на сервер файл базы данных
    keydata = config.RSA_KEY
    key = paramiko.RSAKey(data=decodebytes(keydata))
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys.add(config.host, 'ssh-rsa', key)
    
    con = pysftp.Connection(
    host=config.host,
    username=config.username,
    password=config.password,
    cnopts=cnopts)
    if 'zavtrupoy.db' in con.listdir('/home/web/home/zavtrupoy/'):
        con.remove('/home/web/home/zavtrupoy/zavtrupoy.db')
        con.put('zavtrupoy.db', remotepath='/home/web/home/zavtrupoy/zavtrupoy.db')
    else:
        con.put('zavtrupoy.db', remotepath='/home/web/home/zavtrupoy/zavtrupoy.db')
