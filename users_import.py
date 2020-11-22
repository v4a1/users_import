from bs4 import BeautifulSoup
import csv
from datetime import datetime
from sys import argv


def read_html_file(filename):
    with open(filename, 'r', encoding='utf8') as file:
        data = file.read()
    return data


def read_csv_file(filename, columns_name, codetable, limiter):
    with open(filename, 'r', encoding=codetable) as file:
        lines = csv.DictReader(file, delimiter=limiter)
        data = []
        for line in lines:
            items = {}
            for item in columns_name:
                items.update({item: line[item]})
            data.append(items)
    return data


def save_csv(filename, user_data, columns_name):
    with open(filename, "w", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=columns_name)
        writer.writeheader()
        writer.writerows(user_data)


def pars_activecollab(text):
    soup = BeautifulSoup(text, features="html.parser")
    company_group_list = soup.find('div', {'class': 'company_listing_wrapper'})
    company_groups = company_group_list.find_all(
        'div', {'ng-repeat': 'group in grouped_users | orderBy : orderCompanies track by group.id'})
    all_users = []
    for company_data in company_groups:
        if company_data.find(
                'div', {'class': 'group_listing page_paper'}).find('a', {'ng-if': 'group.id'}) is None:
            company = "Люди без компании"
        else:
            company = company_data.find(
                'div', {'class': 'group_listing page_paper'}).find('a', {'ng-if': 'group.id'}).text
        users_list = company_data.find('tbody')
        items = users_list.find_all(
            'tr', {'ng-repeat': """user in people | orderBy : 'first_name' track by user.id"""})
        for item in items:
            if item.find('td', {'class': "col_name"}).find('a') is not None:
                name = item.find('td', {'class': "col_name"}).find('a').text
                email = item.find(
                    'td', {'class': "col_email"}).find('span').text
                status = item.find(
                    'td', {'class': "col_user_role"}).find('div').text
                user = {"email": email, "name": name,
                        "status": status, "Компания": company}
                all_users.append(user)
    return all_users


def pars_notion(text):
    soup = BeautifulSoup(text, features="html.parser")
    user_list = soup.find('tbody')
    items = user_list.find_all('tr')
    all_users = []
    for item in items:
        name = item.find('div', {'class': "notranslate"}).text
        email = item.find('div', {
                          'style': """font-size: 12px; line-height: 16px; color: rgba(55, 53, 47, 0.6); white-space: nowrap; overflow: hidden; text-overflow: ellipsis;"""}).text
        role = item.find('span', {'class': 'notranslate'}).text
        user = {"email": email, "name": name, "Access level": role}
        all_users.append(user)
    return all_users


def pars_miro(text):
    soup = BeautifulSoup(text, features="html.parser")
    user_list = soup.find('div', {'class': 'company-members-list__content'})
    items = user_list.find_all('div', {'class': "company-member"})
    all_users = []
    for item in items:
        name = item.find('div', {
                         'class': 'company-member--name company-member--ellipsis'}).find('strong').text
        email = item.find(
            'div', {'class': 'company-member--email company-member--ellipsis'}).text
        email = email.replace(' ', '')
        email = email.replace('\n', '')
        email = email.replace('\t', '')
        role = item.find(
            'span', {'ng-bind': 'userWrapper.loadedItem.roleText'}).text
        user = {"email": email, "name": name, "Access level": role}
        all_users.append(user)
    return all_users


def add_data(data_old, data_new, mail, columns):
    for item in data_new:
        last = data_old.get(item.get(mail[1]))
        if last is None:
            last = {}
        for column in columns:
            item_data = item.get(column[1])
            if item_data == "" and (column[1] == "Продукты группы" or column[1] == "Лицензии"):
                item_data = "No license"
            last.update({column[0]: item_data})
        user = {item.get(mail[1]): last}
        data_old.update(user)
    return data_old


def render_sverka(data, columns_name):
    day = datetime.now().strftime("%d.%m.%Y")
    items = [{"email": "Аккаунты", 'Google': day, 'Slack': day, 'ActiveCollab': day, 'Adobe': day,
              'Office365': day, 'Miro': day, 'Notion': day, 'Boomstream': day, 'Jira': day, 'Rarus': day}]
    data_keys = list(data)
    data_keys.sort()
    print(data_keys)
    for user in data_keys:
        item = {}
        for column in columns_name:
            item_temp = data.get(user).get(column)
            if item_temp is None:
                item_temp = "Нет"
            item.update({column: item_temp})
        item.update({columns_name[0]: user})
        # print(item)
        items.append(item)
    return items


def render_tver(data, columns_name):
    day = datetime.now().strftime("%d.%m.%Y")
    items = [{"email": "Почта", "Accounts": day}]
    data_keys = list(data)
    data_keys.sort()
    for user in data_keys:
        item = {}
        user_email = None
        if data.get(user).get("Recovery Email") is not None:
            user_email = data.get(user).get("Recovery Email")
        if data.get(user).get("Home Secondary Email") is not None:
            user_email = data.get(user).get("Home Secondary Email")
        if data.get(user).get("Work Secondary Email") is not None:
            user_email = data.get(user).get("Work Secondary Email")
        if (user_email is not None) and (user_email != "") and (data.get(user).get("Google") != "Suspended"):
            if (data.get(user).get("Slack") is not None) and (data.get(user).get("Slack") != "Deactivated"):
                accounts = "{} + {}".format(user, "Slack")
            if data.get(user).get("ActiveCollab") is not None:
                accounts = "{} + {}".format(accounts, "ActiveCollab")
            item = {"email": user_email, "Accounts": accounts}
            items.append(item)
    return items


script_name, script_argv = argv
if script_argv == "sverka":
    google_data = read_csv_file(
        "google.csv", ["Email Address [Required]", "Status [READ ONLY]"], 'utf8', ",")
    slack_data = read_csv_file("slack.csv", ["email", "status"], 'utf8', ",")
    activecollab_data = pars_activecollab(read_html_file("activecollab.html"))
    adobe_data = read_csv_file(
        "adobe.csv", ["\ufeffЭлектронная почта", "Продукты группы"], 'utf8', ",")
    office365_data = read_csv_file(
        "office365.csv", ["Имя участника-пользователя", "Лицензии"], 'utf8', ",")
    miro_data = pars_miro(read_html_file("miro.html"))
    notion_data = pars_notion(read_html_file("notion.html"))
    boomstream_data = read_csv_file(
        "boomstream.csv", ["Email", "Роль"], 'utf8', ",")
    jira_data = read_csv_file("jira.csv", ["email", "active"], 'utf8', ",")
    rarus_data = read_csv_file("rarus.csv", ["email", "id"], 'utf8', ",")
    data = {}
    data = add_data(data, google_data, ["email", "Email Address [Required]"], [[
                    "Google", "Status [READ ONLY]"]])
    data = add_data(data, slack_data, [
                    "email", "email"], [["Slack", "status"]])
    data = add_data(data, activecollab_data, ["email", "email"], [
                    ["ActiveCollab", "status"]])
    data = add_data(data, adobe_data, ["email", "\ufeffЭлектронная почта"], [
                    ["Adobe", "Продукты группы"]])
    data = add_data(data, office365_data, ["email", "Имя участника-пользователя"], [
        ["Office365", "Лицензии"]])
    data = add_data(data, miro_data, ["email", "email"], [
        ["Miro", "Access level"]])
    data = add_data(data, notion_data, ["email", "email"], [
        ["Notion", "Access level"]])
    data = add_data(data, boomstream_data, ["email", "Email"], [
        ["Boomstream", "Роль"]])
    data = add_data(data, jira_data, ["email", "email"], [
        ["Jira", "active"]])
    data = add_data(data, rarus_data, ["email", "email"], [["Rarus", "id"]])
    columns_name = ["email", 'Google', 'Slack', 'ActiveCollab', 'Adobe',
                    'Office365', 'Miro', 'Notion', 'Boomstream', 'Jira', 'Rarus']
    all_users = render_sverka(data, columns_name)
    save_csv("Данные.csv", all_users, columns_name)

if script_argv == "tver":
    google_data = read_csv_file(
        "google.csv", ["Email Address [Required]", "Status [READ ONLY]", "Recovery Email",
                       "Home Secondary Email", "Work Secondary Email"], 'utf8', ",")
    slack_data = read_csv_file("slack.csv", ["email", "status"], 'utf8', ",")
    activecollab_data = pars_activecollab(read_html_file("activecollab.html"))
    data = {}
    data = add_data(data, google_data, ["email", "Email Address [Required]"], [[
                    "Google", "Status [READ ONLY]"], ["Recovery Email", "Recovery Email"], [
        "Home Secondary Email", "Home Secondary Email"], ["Work Secondary Email",
                                                          "Work Secondary Email"]])
    data = add_data(data, slack_data, ["email", "email"], [
                    ["Slack", "status"]])
    data = add_data(data, activecollab_data, ["email", "email"], [
                    ["ActiveCollab", "status"]])
    columns_name = ["email", 'Accounts']
    all_users = render_tver(data, columns_name)
    save_csv("Тверь.csv", all_users, columns_name)
