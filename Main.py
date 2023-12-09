import traceback
from email.mime.base import MIMEBase
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import smtplib
import schedule
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import yaml
import io
from datetime import datetime
from email import encoders
import uuid
import os


def send_email_with_attachment(sender_email, receiver_email, subject, body, attachment_path, smtp_server,
                               smtp_port, username, password):
    # Create a multipart message
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

    # Add body to the email
    message.attach(MIMEText(body, "plain"))

    # Open the file in binary mode
    with open(attachment_path, "rb") as attachment_file:
        # Create a MIMEBase object and set the appropriate content type
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment_file.read())

    # Encode the attachment in base64
    encoders.encode_base64(part)

    # Set the filename in the Content-Disposition header
    part.add_header("Content-Disposition", f"attachment; filename= {attachment_path}")

    # Attach the file to the email
    message.attach(part)

    # Connect to the SMTP server
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(username, password)
        server.sendmail(sender_email, receiver_email, message.as_string())


def check_shipping_dates():
    try:
        with io.open('config.yaml', encoding='utf-8') as f:
            data = yaml.safe_load(f)
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        # options.add_argument('--headless')
        browser = webdriver.Chrome(options=options)
        wait = WebDriverWait(browser, 500)
        can_be_booked = 0
        smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
        smtpObj.starttls()
        smtpObj.login(data["email"], data["email_password"])
        ready_invoices = open("ready_invoices.txt", "a+", encoding="utf-8")
        ready_invoices.seek(0)

        # Open the Lamoda website
        browser.get("https://gm.lamoda.ru/")
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div/div/div[2]/div/div/div/div/form/div[2]/div/div/div[1]/div/input")))

        # Find the username and password input fields and enter your credentials
        username_input = browser.find_element(By.XPATH,
                                              "/html/body/div/div/div[2]/div/div/div/div/form/div[2]/div/div/div["
                                              "1]/div/input")
        username_input.send_keys(data["username"])
        password_input = browser.find_element(By.XPATH,
                                              "/html/body/div/div/div[2]/div/div/div/div/form/div[4]/div/div/div["
                                              "1]/div/input")
        password_input.send_keys(data["password"])
        browser.find_element(By.XPATH, "/html/body/div/div/div[2]/div/div/div/div/form/div[5]/button").click()

        # Check amount of delivery
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/nav/ul/li[2]/div")))
        browser.find_element(By.XPATH, "/html/body/div[1]/div/nav/ul/li[2]/div").click()
        time.sleep(2)
        print(datetime.now().strftime("%H:%M"), 'Выполняется поиск заявки на бронирование')
        lines = browser.find_elements(By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div[2]/div[2]/table/tbody/tr")
        invoice_list = []
        for line in lines:
            ready_invoices.seek(0)
            ready_invoices_list = ready_invoices.read().split('\n')
            if (line.find_element(By.XPATH, "td[6]").text.split('\n')[0] == "Ожидает брони") and (datetime.strptime(line.find_element(By.XPATH, "td[4]").text, "%d.%m.%Y") >= datetime.strptime(data["start_date"], "%d.%m.%Y")) and (line.find_element(By.XPATH, "td[3]").text not in ready_invoices_list):
                invoice_list.append([line.find_element(By.XPATH, "td[3]").text, int(line.find_element(By.XPATH, "td[5]").text)])
        if invoice_list == []:
            print(datetime.now().strftime("%H:%M"), f"Нет активных заявок")
        for invoice in invoice_list:
            print(datetime.now().strftime("%H:%M"), f'Найдена свободная заявка. Номер заявки: {invoice[0]}. Количество товара: {invoice[1]}')
            browser.find_element(By.XPATH, "/html/body/div[1]/div/nav/ul/li[3]/div").click()
            time.sleep(3)
            print(datetime.now().strftime("%H:%M"), f'Выполняется поиск свободных слотов')
            selected_date = [0, 0]
            free_dates = browser.find_elements(By.CLASS_NAME, "vuecal__event-title-free")
            dates = [[int(i.find_element(By.XPATH, '../..').get_attribute('aria-label')), i] for i in free_dates]
            for n in range(len(dates)):
                if int(dates[n][0]) > int(selected_date[0]) and int(selected_date[0]) != 0 and data["check_only"] == 0:
                    break
                elif int(dates[n][0]) == int(selected_date[0]) and int(dates[n][1].text.split()[-1]) >= invoice[1] and data[
                    "check_only"] == 0:
                    selected_date = dates[n]
                elif int(dates[n][1].text.split()[-1]) >= invoice[1] and data["check_only"] == 0:
                    selected_date = dates[n]
                    can_be_booked = 1
                elif int(dates[n][1].text.split()[-1]) > invoice[1] and data["check_only"] == 1:
                    # msg = MIMEText(
                    #     "Дата: " + str(dates[n][0]) + ', Время: ' + dates[n][1].text.split(', ')[
                    #         0] + ", Доступное количество:" +
                    #     dates[n][1].text.split()[-1])
                    # msg["Subject"] = 'Lbot - есть полный слот'
                    # smtpObj.sendmail(data["email"], data["email_to"], msg.as_string())
                    fullbook_screenshot_name = 'Screenshots/' + str(uuid.uuid4()) + ".png"
                    browser.save_screenshot(fullbook_screenshot_name)

                    send_email_with_attachment(
                        sender_email=data["email"],
                        receiver_email=data["email_to"],
                        subject="LBot - есть полный слот",
                        body='',
                        attachment_path=fullbook_screenshot_name,
                        smtp_server="smtp.gmail.com",
                        smtp_port=587,
                        username=data["email"],
                        password=data["email_password"]
                    )
                    print(datetime.now().strftime("%H:%M"), "LBot - есть полный слот")
                    break
                elif int(dates[n][1].text.split()[-1]) > (invoice[1] // 2) and data["check_only"] == 1:
                    # msg = MIMEText(
                    #     "Дата: " + str(dates[n][0]) + ', Время: ' + dates[n][1].text.split(', ')[
                    #         0] + ", Доступное количество:" +
                    #     dates[n][1].text.split()[-1])
                    # msg["Subject"] = 'Lbot - есть частичный слот'
                    # smtpObj.sendmail(data["email"], data["email_to"], msg.as_string())
                    partbook_screenshot_name = 'Screenshots/' + str(uuid.uuid4()) + ".png"
                    browser.save_screenshot(partbook_screenshot_name)

                    send_email_with_attachment(
                        sender_email=data["email"],
                        receiver_email=data["email_to"],
                        subject="LBot - есть частичный слот",
                        body='',
                        attachment_path=partbook_screenshot_name,
                        smtp_server="smtp.gmail.com",
                        smtp_port=587,
                        username=data["email"],
                        password=data["email_password"]
                    )
                    print(datetime.now().strftime("%H:%M"), "LBot - есть частичный слот")
                    break

            if can_be_booked == 1 and data["check_only"] == 0:
                # selected_day = str(selected_date[0])
                # selected_time = selected_date[1].text.split(', ')[0]
                selected_date[1].click()
                wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/form/div[2]/div[2]/div/div/div/div[2]/div/input")))
                # ТИП СОТРУДНИЧЕСТВА
                browser.find_element(By.XPATH,
                                     "/html/body/div[1]/div/div[2]/div/div/div/form/div[2]/div[2]/div/div/div/div[2]/div/input").click()
                wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[7]/div/div/div[1]/ul/li")))
                browser.find_element(By.XPATH, "/html/body/div[2]/div[7]/div/div/div[1]/ul/li").click()

                # ВЫБОР ПОСТАВКИ
                browser.find_element(By.XPATH,
                                     "/html/body/div[1]/div/div[2]/div/div/div/form/div[2]/div[3]/div[1]/div/div["
                                     "2]/div/div/div/div[2]/div/input").click()
                wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[8]/div/div/div[1]/ul/li")))
                shipments = browser.find_elements(By.XPATH, "/html/body/div[2]/div[8]/div/div/div[1]/ul/li")
                for shipment in shipments:
                    if shipment.text == invoice[0]:
                        shipment.click()

                # КОЛИЧЕСТВО КОРОБОВ
                browser.find_element(By.XPATH,
                                     "/html/body/div[1]/div/div[2]/div/div/div/form/div[2]/div[4]/div[3]/div/div[2]/div/input").send_keys(
                    data["boxes"])

                # ИМЯ ВОДИТЕЛЯ
                browser.find_element(By.XPATH,
                                     "/html/body/div[1]/div/div[2]/div/div/div/form/div[2]/div[6]/div[1]/div/div[2]/div/input").send_keys(
                    data["driver"])

                # НОМЕР МАШИНЫ
                browser.find_element(By.XPATH,
                                     "/html/body/div[1]/div/div[2]/div/div/div/form/div[2]/div[6]/div[2]/div/div[2]/div/input").send_keys(
                    data["car_num"])

                # КНОПКА ЛЕГКОВУШКИ
                browser.find_element(By.XPATH,
                                     "/html/body/div[1]/div/div[2]/div/div/div/form/div[2]/div[7]/div/div/div[2]/label[2]/span[1]").click()

                # НОМЕР ВОДИТЕЛЯ
                browser.find_element(By.XPATH,
                                     "/html/body/div[1]/div/div[2]/div/div/div/form/div[2]/div[9]/div/div/div/input").send_keys(
                    data["driver_phone_num"])

                # ЗАБРОНИРОВАТЬ
                browser.find_element(By.XPATH,
                                     "/html/body/div[1]/div/div[2]/div/div/div/form/div[2]/div[15]/div[2]/div/button").click()
                time.sleep(2)
                if browser.find_element(By.XPATH,
                                     "/html/body/div[1]/div/div[2]/div/div/div/form/div[2]/div[15]/div[2]/span"):
                    browser.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/form/div[2]/div[15]/div[2]/span").click()
                    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/input")))
                    browser.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/input").click()
                    avaliable_dates = browser.find_elements(By.CLASS_NAME, "avaliable")
                    if avaliable_dates:
                        avaliable_dates[0].click()
                        browser.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div[2]/div/div/div/div/div[2]/div[2]/div/div/div/div/div/input").click()
                        avaliable_time = browser.find_elements(By.XPATH, "/html/body/div[2]/div[12]/div/div/div[1]/ul")
                        avaliable_time[-1].click()
                        browser.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div[2]/div/div/footer/span/button[2]").click()
                        time.sleep(2)
                        if browser.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div[2]/div/div/footer/span/button[2]"):
                            bookerror_screenshot_name = 'Screenshots/' + str(uuid.uuid4()) + ".png"
                            browser.save_screenshot(bookerror_screenshot_name)

                            # text = "Дата: " + selected_day + ', Время: ' + selected_time
                            send_email_with_attachment(
                                sender_email=data["email"],
                                receiver_email=data["email_to"],
                                subject="LBot - не удалось забронировать слот",
                                body='',
                                attachment_path=bookerror_screenshot_name,
                                smtp_server="smtp.gmail.com",
                                smtp_port=587,
                                username=data["email"],
                                password=data["email_password"]
                            )
                            print(datetime.now().strftime("%H:%M"), "LBot - не удалось забронировать слот")
                        else:
                            browser.find_element(By.XPATH,
                                                 "/html/body/div[1]/div/nav/ul/li[4]/div").click()
                            time.sleep(2)
                            book_screenshot_name = 'Screenshots/' + str(uuid.uuid4()) + ".png"
                            browser.save_screenshot(book_screenshot_name)

                            # text = "Дата: " + selected_day + ', Время: ' + selected_time
                            send_email_with_attachment(
                                sender_email=data["email"],
                                receiver_email=data["email_to"],
                                subject="LBot - слот забронирован",
                                body='',
                                attachment_path=book_screenshot_name,
                                smtp_server="smtp.gmail.com",
                                smtp_port=587,
                                username=data["email"],
                                password=data["email_password"]
                            )
                            print(datetime.now().strftime("%H:%M"), "LBot - слот забронирован")
                            ready_invoices.write(f"{invoice[0]}\n")
                            ready_invoices.seek(0)
                    else:
                        nobook_screenshot_name = 'Screenshots/' + str(uuid.uuid4()) + ".png"
                        browser.save_screenshot(nobook_screenshot_name)

                        # text = "Дата: " + selected_day + ', Время: ' + selected_time
                        send_email_with_attachment(
                            sender_email=data["email"],
                            receiver_email=data["email_to"],
                            subject="LBot - нет свободных слотов",
                            body='',
                            attachment_path=nobook_screenshot_name,
                            smtp_server="smtp.gmail.com",
                            smtp_port=587,
                            username=data["email"],
                            password=data["email_password"]
                        )
                        print(datetime.now().strftime("%H:%M"), "LBot - нет свободных слотов")
                else:
                    # СКРИНШОТ БРОНИ
                    browser.find_element(By.XPATH,
                                         "/html/body/div[1]/div/nav/ul/li[4]/div").click()
                    time.sleep(2)
                    book_screenshot_name = 'Screenshots/' + str(uuid.uuid4()) + ".png"
                    browser.save_screenshot(book_screenshot_name)

                    # text = "Дата: " + selected_day + ', Время: ' + selected_time
                    send_email_with_attachment(
                                sender_email=data["email"],
                                receiver_email=data["email_to"],
                                subject="LBot - слот забронирован",
                                body='',
                                attachment_path=book_screenshot_name,
                                smtp_server="smtp.gmail.com",
                                smtp_port=587,
                                username=data["email"],
                                password=data["email_password"]
                            )
                    print(datetime.now().strftime("%H:%M"), "LBot - слот забронирован")
                    ready_invoices.write(f"{invoice[0]}\n")
                    ready_invoices.seek(0)
            elif selected_date == [0, 0] and data["check_only"] == 0:
                # msg = MIMEText("Свободных слотов не оказалось")
                # msg["Subject"] = 'Lbot - нет свободных слотов'
                # smtpObj.sendmail(data["email"], data["email_to"], msg.as_string())
                nobook_screenshot_name = 'Screenshots/' + str(uuid.uuid4()) + ".png"
                browser.save_screenshot(nobook_screenshot_name)

                # text = "Дата: " + selected_day + ', Время: ' + selected_time
                send_email_with_attachment(
                    sender_email=data["email"],
                    receiver_email=data["email_to"],
                    subject="LBot - нет свободных слотов",
                    body='',
                    attachment_path=nobook_screenshot_name,
                    smtp_server="smtp.gmail.com",
                    smtp_port=587,
                    username=data["email"],
                    password=data["email_password"]
                )
                print(datetime.now().strftime("%H:%M"), "LBot - нет свободных слотов")
    except Exception:
        print(traceback.format_exc())
        print(datetime.now().strftime("%H:%M"), "Что-то пошло не так, проверьте все по инструкции")


def main():
    with io.open('config.yaml', encoding='utf-8') as f:
        data = yaml.safe_load(f)
    schedule.every(data["minutes"]).minutes.do(check_shipping_dates)

    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == '__main__':
    check_shipping_dates()
    main()
