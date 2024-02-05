import os
import sys
import subprocess
import argparse
import traceback
import configparser
from datetime import datetime
from decimal import Decimal
import pandas
import plyer
import time
from kivy.config import Config
from kivy.app import App
from kivymd.uix.gridlayout import MDGridLayout
from new_pdf import MyDocTemplate
from kivy.clock import Clock
from kivy.uix.screenmanager import ScreenManager, Screen
from kivymd.uix.menu import MDDropdownMenu
import mysql.connector
from kivymd.uix.datatables import MDDataTable
from kivymd.uix.button import MDIconButton
from kivymd.uix.pickers import MDDatePicker
from kivymd.uix.toolbar import MDTopAppBar
from kivy.core.window import Window
from kivymd.uix.button import MDRectangleFlatButton
from kivymd.uix.label import MDLabel
from kivymd.app import MDApp
from kivymd.uix.dialog import MDDialog
from kivymd.uix.textfield import MDTextField
from kivymd.uix.button import MDFlatButton
from kivy.logger import Logger


class LoginApp(MDApp):
    def __init__(self, userlist=None, **kwargs):

        config = configparser.ConfigParser()
        config.read('config.ini')

        self.userlist = userlist
        super().__init__(**kwargs)
        self.sqlstatements = SQLStatements(mandant)

    login_successful = False  # Klassenattribut für den Anmeldestatus

    def build(self):
        self.theme_cls.primary_palette = "Blue"

    def on_start(self):
        self.show_login_popup()

    def show_login_popup(self):
        # Benutzername-Eingabefeld
        self.username_field = MDTextField(
            hint_text="Benutzername",
            required=True,
            helper_text_mode="on_error",
            write_tab=False,
        )

        # Passwort-Eingabefeld
        self.password_field = MDTextField(
            hint_text="Passwort",
            required=True,
            helper_text_mode="on_error",
            password=True,
            write_tab=False
        )

        # Abbrechen-Button
        self.cancel_button = MDFlatButton(
            text="Abbrechen",
            on_release=self.cancel_login
        )

        # Anmelde-Button
        self.login_button = MDFlatButton(
            text="Anmelden",
            on_release=self.show_login_result
        )

        # Fehlermeldung-Label
        self.error_label = MDLabel(
            text="",
            theme_text_color="Error",
            halign="center"
        )

        # Popup-Fenster erstellen
        self.dialog = MDDialog(
            type="custom",
            content_cls=self.create_dialog_content(),
            buttons=[self.cancel_button, self.login_button],
            auto_dismiss=False,
            size_hint=(0.4, None)
        )

        # Popup-Fenster anzeigen
        self.dialog.open()

    def create_dialog_content(self):
        content_box = MDGridLayout(spacing="5dp", cols=4)
        content_box.add_widget(self.username_field)
        content_box.add_widget(self.password_field)
        content_box.add_widget(self.error_label)

        return content_box

    def show_login_result(self, instance):
        username = self.username_field.text
        password = self.password_field.text

        # Überprüfen der Anmeldeinformationen
        for user_data in self.userlist:
            if user_data[1] == username and user_data[3] == password:
                print("Anmeldung erfolgreich!")
                self.dialog.dismiss()  # Popup-Fenster schließen
                self.login_successful = True  # Anmeldestatus auf True setzen
                self.sqlstatements.update_user(username)
                self.stop()

                # Starte die Hauptanwendung (MyMainApp)
                # my_main_app = MyAgrar()
                # my_main_app.run()

        self.error_label.text = "Ungültige Anmeldeinformationen!"

    def cancel_login(self, instance):
        self.dialog.dismiss()  # Popup-Fenster schließen


class SQLStatements():
    def __init__(self, mandant):
        # Verbindung zur Agrar-Datenbank herstellen
        self.role = None
        self.dropdown_table = None
        self.mandant = mandant
        config = configparser.ConfigParser()
        config.read('config.ini')

        self.connection_costumer = mysql.connector.connect(
            host=config[mandant]['HOST'],
            user=config[mandant]['USER'],
            password=config[mandant]['PASSWORD'],
            database=config[mandant]['DBNAME']
        )

        self.connection_general = mysql.connector.connect(
            host=config['GENERAL']['HOST'],
            user=config['GENERAL']['USER'],
            password=config['GENERAL']['PASSWORD'],
            database=config['GENERAL']['DBNAME']
        )

        self.connection_dropdown = mysql.connector.connect(
            host=config['DROPDOWN']['HOST'],
            user=config['DROPDOWN']['USER'],
            password=config['DROPDOWN']['PASSWORD'],
            database=config['DROPDOWN']['DBNAME']
        )

        userlist = self.get_userlist()
        windows_username = os.environ.get('USERNAME')
        for user in userlist:
            if user[4] == windows_username:
                self.role = user[2]
                self.company = user[5]
                break

    def get_services(self):
        # Standard-Tabelle für DLN, diese wird in Abhängigkeit zum User geladen

        if self.role == 'admin':  # volle Sicht auf alle Daten
            cursor = self.connection_costumer.cursor()
            query = f"""SELECT `table_id`, `type`, `description`, DATE_FORMAT(`start_date`, '%d.%m.%Y'), DATE_FORMAT(
            `end_date`, '%d.%m.%Y'), `consulter`, `customer_name`, `status_type_a`, DATE_FORMAT(`timestamp_change`, 
            '%d.%m.%Y') FROM `dienstleistungen` WHERE `company_code` = "{self.company}" ORDER BY `timestamp_change` DESC """
            cursor.execute(query)
            rows = cursor.fetchall()
            # Daten durchgehen und None-Felder ersetzen
            data = [[col if col is not None else "" for col in row] for row in rows]

            cursor.close()
            return data

        if self.role == 'backoffice':  # eingeschränkte Sicht auf alle Daten (keine bezahlten DLN)
            cursor = self.connection_costumer.cursor()
            query = f"""SELECT `table_id`, `type`, `description`, DATE_FORMAT(`start_date`, '%d.%m.%Y'), DATE_FORMAT(
            `end_date`, '%d.%m.%Y'), `consulter`, `customer_name`, `status_type_a`, DATE_FORMAT(`timestamp_change`, 
            '%d.%m.%Y') FROM `dienstleistungen` WHERE `status_type_a` <> "bezahlt" 
            AND `company_code` = {self.company}
            ORDER BY `timestamp_change` DESC"""
            cursor.execute(query)
            rows = cursor.fetchall()
            # Daten durchgehen und None-Felder ersetzen
            data = [[col if col is not None else "" for col in row] for row in rows]

            cursor.close()
            return data

        if self.role == 'beratung' or self.role == 'bewertung':  # eingeschränkte Sicht auf alle Daten (keine
            # bezahlten DLN und keine abgerechneten DLN)
            cursor = self.connection_costumer.cursor()
            query = f"""SELECT `table_id`, `type`, `description`, DATE_FORMAT(`start_date`, '%d.%m.%Y'), DATE_FORMAT( 
            `end_date`, '%d.%m.%Y'), `consulter`, `customer_name`, `status_type_a`, DATE_FORMAT(`timestamp_change`, 
            '%d.%m.%Y') FROM `dienstleistungen` WHERE `status_type_a` <> "bezahlt" 
            AND `status_type_a` <> "abgerechnet" AND `company_code` = {self.company}
            ORDER BY `timestamp_change` DESC """
            cursor.execute(query)
            rows = cursor.fetchall()
            # Daten durchgehen und None-Felder ersetzen
            data = [[col if col is not None else "" for col in row] for row in rows]

            cursor.close()
            return data

    def get_services_with_id(self, selected_service):
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT `table_id`, `description`, DATE_FORMAT(`start_date`, '%d.%m.%Y'), DATE_FORMAT( 
        `end_date`, '%d.%m.%Y'), `customer_name`, `consulter`, `type`, `last_user`, `abrechnungstyp`, `abrechnungszyklus` FROM `dienstleistungen` WHERE 
        `table_id` = {selected_service[0]} """
        # AND `company_code` = {self.company}
        # WHERE `deletion_flag` <> "True" AND status_type_a <> "abgerechnet" AND status_type_a <> "bezahlt"
        #         ORDER BY `table_id` DESC
        cursor.execute(query)
        rows = cursor.fetchall()
        # Daten durchgehen und None-Felder ersetzen
        data = [[col if col is not None else "" for col in row] for row in rows]
        cursor.close()
        return data

    def get_adress(self, selected_service):
        cursor = self.connection_general.cursor()
        query = f"""SELECT * FROM `kunden` WHERE `customer_name` = "{selected_service[6]}" """
        cursor.execute(query)
        data = cursor.fetchall()
        cursor.close()
        return data

    def get_positions(self, selected_service):
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT `table_id`, `type`, `description`, `comment`, DATE_FORMAT(`start_date`, '%d.%m.%Y'),
        `amount`, `value_a`, `value`, `consulter`, `group_a` FROM `positionen` WHERE `task_number` = "{selected_service[0]}"
                """
        cursor.execute(query)
        data_positions = cursor.fetchall()
        return data_positions

    def get_servicetypes(self):
        cursor = self.connection_dropdown.cursor()
        query = f"""SELECT * FROM `001` WHERE `Typ` = "Dienstleistung" AND `Company` = '{self.company}' """
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result

    def get_servicetypes_from_db(self):
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT DISTINCT `type` FROM `dienstleistungen` """
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result

    def get_servicedunning_man_from_db(self):
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT DISTINCT `status_type_c` FROM `dienstleistungen` """
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result

    def get_servicedunning_auto_from_db(self):
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT DISTINCT `status_type_b` FROM `dienstleistungen` """
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result

    def get_consulter(self, text):
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT `lastname` FROM `nutzer` WHERE `lastname` LIKE "%{text}%" """
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result, text

    def get_consulter_from_db(self):
        # es werden alle consulter gesucht, die schon mal was angelegt haben
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT DISTINCT `consulter` FROM `dienstleistungen` """
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result

    def get_consulter_from_db_1(self):
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT DISTINCT `consulter` FROM `dienstleistungen` WHERE `status_type_a` = "abgerechnet" OR 
        `status_type_a` = "bezahlt" """
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result

    def get_customer(self, text):
        # Müsste eigentlich für jedes Unternehmen einzeln sein,
        # Kundentabelle ist angelegt aber leer
        cursor = self.connection_general.cursor()
        query = f""" SELECT `customer_id`,`customer_name` FROM `kunden` WHERE `customer_name` LIKE "%{text}%" """
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result, text

    def get_responsible(self, text):
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT `table_id`,`lastname` FROM `nutzer` WHERE `group_a` = {1} """
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result, text

    def get_customer_from_db(self):
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT DISTINCT `customer_name` FROM `dienstleistungen` """
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result

    def get_customer_from_db_1(self):
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT DISTINCT `customer_name` FROM `dienstleistungen`WHERE `status_type_a` = "abgerechnet" OR 
        `status_type_a` = "bezahlt" """
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result

    def get_status_from_db(self):
        cursor = self.connection_costumer.cursor()
        if self.role == 'admin':
            query = f""" SELECT DISTINCT `status_type_a` FROM `dienstleistungen` """
        if self.role == 'backoffice':
            query = f""" SELECT DISTINCT `status_type_a` FROM `dienstleistungen` WHERE `status_type_a`<> "bezahlt" """
        if self.role == 'beratung' or self.role == 'bewertung':
            query = f""" SELECT DISTINCT `status_type_a` FROM `dienstleistungen` WHERE `status_type_a`<> "bezahlt" 
                AND `status_type_a`<> "abgerechnet" """
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result

    def filter_datatable(self, attribute_list):
        table_name = "`dienstleistungen`"
        cursor = self.connection_costumer.cursor()

        # Überprüft, ob die erforderlichen Spalten in attribute_list enthalten sind
        required_columns = ["table_id", "type", "description", "DATE_FORMAT(start_date, '%d.%m.%Y')",
                            "DATE_FORMAT(end_date, '%d.%m.%Y')", "consulter", "customer_name",
                            "status_type_a", "DATE_FORMAT(timestamp_change, '%d.%m.%Y')"]
        for column in required_columns:
            if column not in attribute_list:
                attribute_list[column] = ""

        # Erstellt eine separate Liste für die SELECT-Klausel, um die Reihenfolge der Spalten beizubehalten
        select_columns = [column for column in required_columns if column in attribute_list]

        # Erstellt die SELECT-Klausel mit den übergebenen Attributen in der gewünschten Reihenfolge
        select_clause = "SELECT " + ", ".join(select_columns)

        # Erstellt die WHERE-Klausel mit den Bedingungen für nicht leere Felder in den entsprechenden Spalten
        where_conditions = []
        for column, attribute in attribute_list.items():
            if attribute != "..." and attribute != "":  # Überprüft, ob das Attribut nicht leer ist
                where_conditions.append(column + " = '" + attribute + "'")
        if where_conditions:  # Überprüft, ob es nicht leere Bedingungen gibt
            where_clause = "WHERE " + " AND ".join(where_conditions)
        else:
            where_clause = ""  # Wenn alle Bedingungen leer sind, wird keine WHERE-Klausel verwendet

        # Fügt die Sortierung nach timestamp_create in absteigender Reihenfolge hinzu
        order_by_clause = "ORDER BY timestamp_create DESC"

        # Kombiniert die SELECT- und WHERE-Klausel zu einer vollständigen SQL-Abfrage
        sql_query = select_clause + "\nFROM " + table_name + "\n" + where_clause + "\n" + order_by_clause
        # Führt die Abfrage mit dem Cursor aus
        cursor.execute(sql_query)

        # Abrufen der Ergebnisse, z.B. als Liste von Tupeln
        results = cursor.fetchall()

        # Ersetzt None-Werte in den Ergebnissen durch ""
        results = [[value if value is not None else "" for value in row] for row in results]

        return results

    def filter_datatable_1(self, attribute_list):
        table_name = "`dienstleistungen`"
        cursor = self.connection_costumer.cursor()

        # Überprüft, ob die erforderlichen Spalten in attribute_list enthalten sind
        required_columns = ["table_id", "DATE_FORMAT(`timestamp_accounting`, '%d.%m.%Y')",
                            "DATE_FORMAT(`timestamp_payment`, '%d.%m.%Y')", "description",
                            "consulter", "customer_name", "status_type_a", "status_type_b", "status_type_c"]
        for column in required_columns:
            if column not in attribute_list:
                attribute_list[column] = ""

        # Erstellt eine separate Liste für die SELECT-Klausel, um die Reihenfolge der Spalten beizubehalten
        select_columns = [column for column in required_columns if column in attribute_list]

        # Erstellt die SELECT-Klausel mit den übergebenen Attributen in der gewünschten Reihenfolge
        select_clause = "SELECT " + ", ".join(select_columns)

        # Erstellt die WHERE-Klausel mit den Bedingungen für nicht leere Felder in den entsprechenden Spalten
        where_conditions = []
        for column, attribute in attribute_list.items():
            if attribute != "..." and attribute != "":  # Überprüft, ob das Attribut nicht leer ist
                where_conditions.append(column + " = '" + attribute + "'")

        if where_conditions:  # Überprüft, ob es nicht leere Bedingungen gibt
            where_clause = "WHERE " + " AND ".join(where_conditions)
        else:
            where_clause = ""  # Wenn alle Bedingungen leer sind, wird keine WHERE-Klausel verwendet

        # Fügt die Sortierung nach timestamp_create in absteigender Reihenfolge hinzu
        # order_by_clause = "ORDER BY CASE WHEN `status_type_b` = 'Inkasso' THEN 1 WHEN `status_type_b` = '3. Mahnung' THEN 2 WHEN `status_type_b` = '2. Mahnung' THEN 3 WHEN `status_type_b` = '1. Mahnung' THEN 4 ELSE 5"

        # Kombiniert die SELECT- und WHERE-Klausel zu einer vollständigen SQL-Abfrage
        sql_query = select_clause + "\nFROM " + table_name + "\n" + where_clause + "\n"  # + order_by_clause
        # Führt die Abfrage mit dem Cursor aus
        cursor.execute(sql_query)

        # Abrufen der Ergebnisse, z.B. als Liste von Tupeln
        results = cursor.fetchall()

        # Ersetzt None-Werte in den Ergebnissen durch ""
        results = [[value if value is not None else "" for value in row] for row in results]

        return results

    def input_service(self, service_type, description, consulter, customer, area, abrechnungstyp, abrechnungszyklus):
        cursor = self.connection_costumer.cursor()
        windows_username = os.environ.get('USERNAME')
        timestamp = datetime.now()
        # Dienstleistungen werden initial immer mit dem Status "angelegt" angelegt
        # Es wird ein timestamp und der anlegende User hinterlegt
        query = "INSERT INTO `dienstleistungen` (timestamp_create, description, last_user, status_type_a, " \
                "customer_name, customer_id, consulter, type, group_b, company_code, abrechnungstyp, abrechnungszyklus ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
        values = (
            timestamp, description, windows_username, "angelegt", customer, "", consulter, service_type, area,
            self.company, abrechnungstyp, abrechnungszyklus
        )
        cursor.execute(query, values)
        self.connection_costumer.commit()
        cursor.close()

    def update_service(self, service_type, description, consulter, customer, list, abrechnungstyp, abrechnungszyklus):
        cursor = self.connection_costumer.cursor()
        windows_username = os.environ.get('USERNAME')
        timestamp = datetime.now()
        # Hier können Sie den Code zum Ausführen von SQL-Abfragen oder zum Einfügen von Daten schreiben
        query = "UPDATE `dienstleistungen` SET timestamp_change = %s, description = %s, last_user = %s, customer_name = " \
                "%s, consulter = %s, type = %s, abrechnungstyp = %s, abrechnungszyklus = %s WHERE table_id = %s"
        values = (timestamp, description, windows_username, customer, consulter, service_type, abrechnungstyp, abrechnungszyklus , list[0])
        cursor.execute(query, values)
        self.connection_costumer.commit()
        cursor.close()

    def update_service_header(self, selected_service, min_date, max_date):
        cursor = self.connection_costumer.cursor()
        windows_username = os.environ.get('USERNAME')
        timestamp = datetime.now()

        # date_obj_1 = datetime.strptime(min_date, "%d.%m.%Y")
        # date_obj_2 = datetime.strptime(max_date, "%d.%m.%Y")

        # Hier können Sie den Code zum Ausführen von SQL-Abfragen oder zum Einfügen von Daten schreiben
        query = "UPDATE `dienstleistungen` SET `start_date` = %s, `end_date` = %s, `last_user` = %s WHERE `table_id` " \
                "= %s"
        values = (min_date, max_date, windows_username, selected_service[0])

        cursor.execute(query, values)
        self.connection_costumer.commit()
        cursor.close()

    def get_positiontypes(self, list):
        cursor = self.connection_dropdown.cursor()
        query = f"""SELECT * from `001` WHERE `Beschreibung` = "{list[1]}" """
        cursor.execute(query)
        result = cursor.fetchall()
        # Tabelle gefunden, wo die Dropdown Item drin stehen
        self.dropdown_table = result[0][1]
        cursor = self.connection_dropdown.cursor()
        query = f"""SELECT `Beschreibung` from {self.dropdown_table} WHERE `Company`= {self.company}"""
        cursor.execute(query)
        result = cursor.fetchall()
        return result

    def get_positiondescription(self, choice):
        cursor = self.connection_dropdown.cursor()
        print("Verbindung zum Server hergestellt")
        query = f"""SELECT * from {self.dropdown_table} WHERE `Beschreibung` = "{choice}" """
        cursor.execute(query)
        result = cursor.fetchall()
        value = result[0][2]
        print(result)
        self.dropdown_table = result[0][1]
        query = f"""SELECT `Beschreibung`, `Umsatzverteilung` from {self.dropdown_table} """
        cursor.execute(query)
        result = cursor.fetchall()
        return result, value

    def input_and_calculate_positions(self, position_type, position_description, position_datefield, position_amount,
                                      position_singleprice, position_consulter, list, position_comment, responsible):
        try:
            position_amount = position_amount.replace(",", ".")
            print("Umwandlung Menge hat funktioniert")

        except:
            print("Umwandlung Menge hat nicht funktioniert oder war nicht nötig")

        try:
            position_singleprice = position_singleprice.replace(",", ".")
            print("Umwandlung Einzelpreis hat funktioniert")

        except:
            print("Umwandlung Einzelpreis hat nicht funktioniert oder war nicht nötig")

        position_price_sum = float(position_amount) * float(position_singleprice)

        cursor = self.connection_costumer.cursor()
        windows_username = os.environ.get('USERNAME')
        timestamp = datetime.now()

        date_object = datetime.strptime(position_datefield, '%Y-%m-%d')
        # Hier können Sie den Code zum Ausführen von SQL-Abfragen oder zum Einfügen von Daten schreiben
        query = "INSERT INTO `positionen` (timestamp_create, description, last_user, status_type_a, consulter, " \
                "type, start_date, amount, value_a, value, task_number, comment, group_a) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
        values = (timestamp, position_description, windows_username, "angelegt", position_consulter, position_type,
                  date_object, position_amount, position_singleprice, position_price_sum, list[0],
                  position_comment, responsible)
        cursor.execute(query, values)
        self.connection_costumer.commit()
        cursor.close()

    def update_positions(self, position_type, position_description, position_datefield, position_amount,
                         position_singleprice, position_consulter, selected_position, position_comment, responsible):
        try:
            position_amount = position_amount.replace(",", ".")
            print("Umwandlung Menge hat funktioniert")

        except:
            print("Umwandlung Menge hat nicht funktioniert oder war nicht nötig")

        try:
            position_singleprice = position_singleprice.replace(",", ".")
            print("Umwandlung Einzelpreis hat funktioniert")

        except:
            print("Umwandlung Einzelpreis hat nicht funktioniert oder war nicht nötig")

        position_price_sum = float(position_amount) * float(position_singleprice)

        position_amount = Decimal(position_amount)
        position_singleprice = Decimal(position_singleprice)
        position_price_sum = Decimal(position_price_sum)

        cursor = self.connection_costumer.cursor()
        windows_username = os.environ.get('USERNAME')
        try:
            date_object_3 = ""
            date_object_3 = datetime.strptime(position_datefield, '%Y-%m-%d')
        except:
            date_object_3 = ""
            date_object_3 = datetime.strptime(position_datefield, '%d.%m.%Y')

        # Hier können Sie den Code zum Ausführen von SQL-Abfragen oder zum Einfügen von Daten schreiben
        query = "UPDATE `positionen` SET `description` = %s, `last_user` = %s, `status_type_a` = %s, " \
                "`consulter`= %s, `type` = %s, `start_date` = %s, `amount` = %s, `value_a` = %s, `value` = %s, " \
                "`comment`= %s , `group_a` = %s WHERE `table_id` = %s"
        values = (position_description, windows_username, "angelegt", position_consulter, position_type,
                  date_object_3, position_amount, position_singleprice, position_price_sum, position_comment,
                  responsible,
                  str(selected_position[0]))
        cursor.execute(query, values)
        self.connection_costumer.commit()
        cursor.close()

    def get_positiongroups(self, selected_service):
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT DISTINCT `type` FROM `positionen` WHERE `task_number` = {selected_service[0]}"""
        cursor.execute(query)
        positiongroups = cursor.fetchall()
        positiongroups = [element[0] for element in positiongroups]
        return positiongroups

    def get_positions_with_group_and_id(self, selected_service, position):
        cursor = self.connection_costumer.cursor()
        query = f""" SELECT `description`,`comment`, DATE_FORMAT(`start_date`, '%d.%m.%Y'), `amount`, `value_a`, 
        `value`, `consulter` from `positionen` WHERE `task_number` = {selected_service[0]} AND `type` = "{position}" """
        cursor.execute(query)
        position_itmes = cursor.fetchall()
        return position_itmes

    def change_status(self, status, selected_service):
        # es muss geschaut werden welcher Statuswechsel reinkommt
        cursor = self.connection_costumer.cursor()
        windows_username = os.environ.get('USERNAME')
        timestamp = datetime.now()
        if status == "bearbeitet":
            query = f""" UPDATE `dienstleistungen` SET `status_type_a` = %s, `last_user` = %s, `timestamp_change` = 
            %s WHERE `table_id` = %s"""
        if status == "abgerechnet":
            query = f""" UPDATE `dienstleistungen` SET `status_type_a` = %s, `last_user` = %s, `timestamp_accounting` = 
            %s WHERE `table_id` = %s"""
        if status == "bezahlt":
            query = f""" UPDATE `dienstleistungen` SET `status_type_a` = %s, `last_user` = %s, `timestamp_payment` = 
            %s WHERE `table_id` = %s"""
        if status == "abgeschlossen":
            query = f""" UPDATE `dienstleistungen` SET `status_type_a` = %s, `last_user` = %s, `timestamp_closing` = 
                    %s WHERE `table_id` = %s"""
        values = (status, windows_username, timestamp, selected_service[0])
        cursor.execute(query, values)
        self.connection_costumer.commit()
        cursor.close()

    def delete_position(self, selected_position):
        cursor = self.connection_costumer.cursor()
        query = f""" DELETE from `positionen` WHERE `table_id` = {selected_position[0]} """
        cursor.execute(query)
        self.connection_costumer.commit()
        cursor.close()

    def get_userlist(self):
        cursor = self.connection_costumer.cursor()
        query = "SELECT `firstname`, `lastname`, `role`, `password` , `last_user`, `company_code` FROM `nutzer` "
        cursor.execute(query)
        userlist = cursor.fetchall()
        return userlist

    def update_user(self, name):
        cursor = self.connection_costumer.cursor()
        windows_username = os.environ.get('USERNAME')
        query = f""" UPDATE `nutzer` SET `last_user` = %s WHERE `lastname` = %s """
        values = (windows_username, name)
        cursor.execute(query, values)
        self.connection_costumer.commit()
        cursor.close()

    def input_customer(self, name, id, adress, town, postalcode, phonenumber, verantortlich, crm_nummer):
        cursor = self.connection_general.cursor()
        windows_username = os.environ.get('USERNAME')
        timestamp = datetime.now()
        query = "INSERT INTO `kunden` (customer_name, customer_id, customer_adress, customer_town, " \
                "customer_postal, last_user, group_a, group_b, group_c) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s )"
        values = (name, id, adress, town, postalcode, windows_username, phonenumber, verantortlich, crm_nummer)
        cursor.execute(query, values)
        self.connection_general.commit()
        cursor.close()

    def delete_crm_data(self):
        cursor = self.connection_general.cursor()
        query = f""" DELETE from `kunden` """
        cursor.execute(query)
        self.connection_general.commit()
        cursor.close()

    def get_services_for_dunning(self):
        cursor = self.connection_costumer.cursor()
        query = f"""
            SELECT `table_id`, DATE_FORMAT(`timestamp_accounting`, '%d.%m.%Y'), DATE_FORMAT(`timestamp_payment`, '%d.%m.%Y'),
            `description`, `consulter`, `customer_name`, `status_type_a`, `status_type_b`, `status_type_c`
            FROM `dienstleistungen`
            WHERE `status_type_a` = 'abgerechnet' OR `status_type_a` = 'bezahlt' AND `company_code` = {self.company}
            ORDER BY 
                CASE
                    WHEN `status_type_b` = 'Inkasso' THEN 1
                    WHEN `status_type_b` = '3. Mahnung' THEN 2
                    WHEN `status_type_b` = '2. Mahnung' THEN 3
                    WHEN `status_type_b` = '1. Mahnung' THEN 4
                    ELSE 5
            END
            """

        cursor.execute(query)
        table_data = cursor.fetchall()

        table_data = [[value if value is not None else "" for value in row] for row in table_data]

        return table_data

    def get_dunning_options(self):
        cursor = self.connection_dropdown.cursor()
        query = "SELECT * FROM `_008` "
        # Vielleicht müsste man hier in 100 suchen und dann nach Mahnungen?
        cursor.execute(query)
        result = cursor.fetchall()
        return result

    def update_service_dunning(self, selected_service, dunning_stage):
        # Mahnstufe wird manuell vom User gesetzt (relevant)
        cursor = self.connection_costumer.cursor()
        windows_username = os.environ.get('USERNAME')
        query = f""" UPDATE `dienstleistungen` SET `last_user` = %s, `status_type_c`= %s WHERE `table_id` = %s """
        values = (windows_username, dunning_stage, selected_service[0])
        cursor.execute(query, values)
        self.connection_costumer.commit()
        cursor.close()

    def update_service_dunning_auto(self, service, auto_dunning, man_dunning):
        # Mahnstufe wird automatisch vom System gesetzt (relevant)
        cursor = self.connection_costumer.cursor()
        windows_username = os.environ.get('USERNAME')
        query = f""" UPDATE `dienstleistungen` SET `last_user` = %s, `status_type_b`= %s , `status_type_c`= %s WHERE `table_id` = %s """
        values = (windows_username, auto_dunning, man_dunning, service[0])
        cursor.execute(query, values)
        self.connection_costumer.commit()
        cursor.close()

    def get_multidropdown_dunning_id(self):
        cursor = self.connection_dropdown.cursor()


class ServicesScreen(Screen):
    service_filter: MDLabel

    def __init__(self, **kwargs):
        super(ServicesScreen, self).__init__(**kwargs)
        self.create_position_grid_layout = None
        self.service_view_label = None
        self.position_selected = False
        self.position_info = []
        self.service_info = []
        self.service_selected = False
        self.name = "services"
        self.create_service_dialog = False
        self.create_position_dialog = False
        self.sql_statements = SQLStatements(mandant)
        self.informations = None
        self.versionsnummer = "1.28"
        self.mandant = self.sql_statements.mandant
        self.database = self.sql_statements.connection_costumer.database

        userlist = self.sql_statements.get_userlist()
        windows_username = os.environ.get('USERNAME')
        for user in userlist:
            if user[4] == windows_username:
                self.role = user[2]
                self.username = user[0]
                break

        dropdown = MDIconButton(icon="dots-vertical", pos_hint={"x": 0, "center_y": 0.5})
        dropdown.bind(on_release=self.dropdown_rooms)

        toolbar = MDTopAppBar(
            pos_hint={'top': 1},
            title="",
            md_bg_color=[0.18, 0.19, 0.19, 1],
            left_action_items=[["dots-vertical", lambda x: self.dropdown_rooms(x)],
                               ["format-list-numbered", lambda x: self.number_filter(x), "Nummernfilter"],
                               ["format-list-bulleted-type", lambda x: self.service_type_multi_filter(x),
                                "Dienstleistungstyp auswählen"],
                               ["account-question", lambda x: self.consulter_multi_filter(x),
                                "Ansprechpartner auswählen"],
                               ["account-question-outline", lambda x: self.customer_multi_filter(x),
                                "Rechnungsempfänger auswählen"],
                               ["list-status", lambda x: self.status_multi_filter(x), "Status auswählen"],
                               ["filter-check-outline", lambda x: self.use_multifilter(x), "Filter anwenden"],
                               ["filter-remove-outline", lambda x: self.reset_filter(x), "Filter zurücksetzen"],
                               ],
            right_action_items=[
                ["plus", lambda x: self.create_service(x), "Dienstleistung anlegen"],
                ["file-edit-outline", lambda x: self.edit_service(x), "Dienstleistung bearbeiten"],
                ["file-document-remove-outline", lambda x: self.delete_service(x), "Dienstleistung löschen"],
                ["file-pdf-box", lambda x: self.view_service(x), "Dienstleistung drucken"],
                ["email-arrow-right-outline", lambda x: self.send_service(x), "Dienstleistung senden"],
                ["list-status", lambda x: self.change_status(x), "Status der Dienstleistung ändern"],
                ["account-question-outline", lambda x: self.get_user_data(x), "Wer bin ich"],
                ["", lambda x: "",
                 "", ""],
                ["", lambda x: "",
                 "", ""],
                ["", lambda x: "",
                 "", ""],
                ["playlist-plus", lambda x: self.create_position(x),
                 "Position hinzufügen", ""],
                ["playlist-edit", lambda x: self.edit_positions(x),
                 "Position ändern", ""],
                ["playlist-remove", lambda x: self.delete_position(x),
                 "Position löschen", ""],
            ],
        )
        self.add_widget(toolbar)

    def number_filter(self, instance):
        # dropdown aller Nummer der Services
        result = self.sql_statements.get_services()
        # Liste nach dem ersten Eintrag jeder Position sortieren
        sorted_list = sorted(result, key=lambda x: x[0], reverse=True)
        self.list = [
            {
                "viewclass": "OneLineListItem",
                "text": f"{i[0]}",
                "on_release": lambda x=f"{i[0]}": self.filter_servicetable_with_number(x, sorted_list)
            }
            for i in sorted_list
        ]
        # Dropdown wird technisch erzeugt
        self.dropdown = MDDropdownMenu(
            items=self.list,
            width_mult=4,
            position="auto"
        )
        self.dropdown.caller = instance
        self.dropdown.open()

    def filter_servicetable_with_number(self, instance, sorted_list):
        data = [eintrag for eintrag in sorted_list if eintrag[0] == int(instance)]
        print(data)
        self.service_table.row_data = data
        self.dropdown.dismiss()
        self.service_view_label.text = "..."
        self.service_selected = False

    def update_selected_items_label(self):
        self.selected_items_label.text = "\n".join(self.selected_items)

    def get_user_data(self, instance):
        self.informations.text = "Hallo " + str(self.username) + ", du bist als " + str(self.role) + " angemeldet."
        Clock.schedule_once(self.clear_information, 5)

    def on_enter(self):
        self.filter_grid = MDGridLayout(cols=4, line_color="white", pos_hint={'x': 0.21, 'y': 0.95},
                                        size_hint=(0.45, 0.04), padding=5)

        self.service_type_filter = MDLabel(text='...', theme_text_color='Custom', text_color=(1, 1, 1, 1))
        self.service_consulter_filter = MDLabel(text='...', theme_text_color='Custom', text_color=(1, 1, 1, 1))
        self.service_customer_filter = MDLabel(text='...', theme_text_color='Custom', text_color=(1, 1, 1, 1))
        self.service_status_filter = MDLabel(text='...', theme_text_color='Custom', text_color=(1, 1, 1, 1))

        self.add_widget(self.filter_grid)
        self.filter_grid.add_widget(self.service_type_filter)
        self.filter_grid.add_widget(self.service_consulter_filter)
        self.filter_grid.add_widget(self.service_customer_filter)
        self.filter_grid.add_widget(self.service_status_filter)

        # Daten für die Tabelle abrufen
        data_services = self.sql_statements.get_services()
        data_positions = []

        # Spaltenüberschriften für die Tabelle
        service_column_data = [
            ("Nr.", Window.width * 0.01),
            ("Dienstleistungstyp", Window.width * 0.05),
            ("Bezeichnung", Window.width * 0.05),
            ("Start", Window.width * 0.025),
            ("Ende", Window.width * 0.025),
            ("Ansprechpartner", Window.width * 0.04),
            ("Rechnungsempfänger", Window.width * 0.05),
            ("Status", Window.width * 0.03),
            ("Änderungsdatum", Window.width * 0.035)
        ]

        position_column_data = [
            ("Nr.", Window.width * 0.01),
            ("Dienstleistungstyp", Window.width * 0.05),
            ("Bezeichnung", Window.width * 0.05),
            ("Bemerkung", Window.width * 0.04),
            ("Datum", Window.width * 0.025),
            ("Menge", Window.width * 0.02),
            ("Einzelpreis", Window.width * 0.025),
            ("Positionspreis", Window.width * 0.03),
            ("Ansprechpartner", Window.width * 0.04),
            ("Umsatzverteilung", Window.width * 0.1)
        ]

        # cell_font_size = 13
        # data_s = []
        # for row in data_services:
        #    cells = []
        #    for cell in row:
        #        cells.append(f'[size={cell_font_size}]{cell}[/size]')
        #    data_s.append(cells)

        # MDDataTable erstellen
        self.service_table = MDDataTable(
            size_hint=(0.7, 0.4),  # Größe des Tabellenwidgets im Verhältnis zum Screen
            pos_hint={"x": 0.03, "y": 0.5},
            column_data=service_column_data,
            row_data=data_services,
            use_pagination=False,  # Bei vielen Daten die Verwendung von Seitenumbrüchen ermöglichen (optional)
            pagination_menu_pos="center",  # Position des Seitenumbruchmenüs (optional)
            rows_num=10000,  # Anzahl der Zeilen pro Seite festlegen
        )

        self.position_table = MDDataTable(
            size_hint=(0.7, 0.4),  # Größe des Tabellenwidgets im Verhältnis zum Screen
            pos_hint={"x": 0.03, "y": 0.05},
            column_data=position_column_data,
            row_data=data_positions,
            use_pagination=False,  # Bei vielen Daten die Verwendung von Seitenumbrüchen ermöglichen (optional)
            pagination_menu_pos="auto",  # Position des Seitenumbruchmenüs (optional)
            rows_num=1000,  # Anzahl der Zeilen pro Seite festlegen
        )

        # MDDataTable im Screen anzeigen
        self.service_table.bind(on_row_press=self.on_row_press_service)
        self.position_table.bind(on_row_press=self.on_row_press_positions)
        self.service_view_label = MDLabel(text="...", pos_hint={"x": 0.04, "y": 0.42})
        self.service_view_label.font_size = "20sp"
        self.position_view_label = MDLabel(text="...", pos_hint={"x": 0.04, "y": -0.03})
        self.position_view_label.font_size = "20sp"
        self.informations = MDLabel(text="...", pos_hint={"x": 0.76, "y": 0.42})
        self.informations.font_size = "20sp"
        self.versionsnummer_label = MDLabel(
            text="Version " + str(self.versionsnummer) + " Mandant: " + str(self.mandant) + " Datenbank: " + str(
                self.database),
            pos_hint={"x": 0.04, "y": -0.47})

        self.add_widget(self.service_table)
        self.add_widget(self.position_table)
        self.add_widget(self.service_view_label)
        self.add_widget(self.position_view_label)
        self.add_widget(self.informations)
        self.add_widget(self.versionsnummer_label)

        self.informations.text = "Hallo " + str(self.username) + ", du bist als " + str(self.role) + " angemeldet."
        Clock.schedule_once(self.clear_information, 10)

    def on_row_press_service(self, instance_table, instance_row, *args):
        selected_range = instance_row.range
        tabledata = instance_row.table.row_data

        new_data = [value for sublist in tabledata for value in sublist]
        info = self.service_info
        self.service_info = []

        for index in range(selected_range[0], selected_range[1] + 1):
            self.service_info.append(new_data[index])

        self.service_view_label.text = "Folgender DLN wurde gewählt: " + str(
            self.service_info[1]) + " mit der Nr. " + str(
            self.service_info[0])
        self.service_selected = True

        data_positions = self.sql_statements.get_positions(self.service_info)
        self.position_table.row_data = data_positions

    def on_row_press_positions(self, instance_table, instance_row, *args):
        selected_range = instance_row.range
        tabledata = instance_row.table.row_data
        all_values = [value for sublist in tabledata for value in sublist]
        self.position_info = []

        for index in range(selected_range[0], selected_range[1] + 1):
            self.position_info.append(all_values[index])

        self.position_view_label.text = "Folgende Position wurde gewählt: " + str(
            self.position_info[1]) + " mit der Nr. " + str(
            self.position_info[0])
        self.position_selected = True

        data_positions = self.sql_statements.get_positions(self.service_info)
        self.position_table.row_data = data_positions

    def on_leave(self):
        if self.create_service_dialog:  # Überprüfen, ob der Dialog geöffnet ist
            self.create_service_dialog = False  # Den Dialog als geschlossen markieren
            self.create_position_dialog = False
        # Liste der zu überprüfenden Klassen
        classes_to_remove = (MDDataTable, MDLabel, MDGridLayout)

        # Entfernen der MDDataTable und MDLabel aus dem Widget-Tree
        widgets_to_remove = [child for child in self.children if isinstance(child, classes_to_remove)]
        for widget in widgets_to_remove:
            self.remove_widget(widget)

    def create_service(self, instance):
        if not self.create_service_dialog:  # Überprüfen, ob der Dialog bereits geöffnet ist
            self.create_service_dialog = True  # Den Dialog als geöffnet markieren
            # GridLayout erstellen
            self.create_service_grid_layout = MDGridLayout(cols=1, padding=10, size_hint=(0.2, 0.3),
                                                           pos_hint={'x': 0.75, 'y': 0.6})
            self.create_service_grid_layout.line_width = 1

            self.create_service_button_grid = MDGridLayout(cols=2, padding=10, size_hint=(0.2, 0.05),
                                                           pos_hint={'x': 0.75, 'y': 0.5})
            self.create_service_button_grid.line_width = 1

            # Fehlermeldungstext erstellen
            self.service_error_label = MDLabel(text='Dienstleistung anlegen', size_hint=(0.2, 0.1))
            self.create_service_grid_layout.add_widget(self.service_error_label)

            # Texteingabefelder erstellen
            self.service_type = MDTextField(hint_text='Dienstleistungstyp', write_tab=False)
            self.description = MDTextField(hint_text='Bezeichnung', write_tab=False)
            self.consulter = MDTextField(hint_text='Ansprechpartner', write_tab=False)
            self.customer = MDTextField(hint_text='Rechnungsempfänger', write_tab=False)
            self.abrechnugszyklus = MDTextField(hint_text='Abrechnungszyklus', write_tab=False)
            self.abrechnungstyp = MDTextField(hint_text='Abrechnungstyp', write_tab=False)
            # self.responsible = MDTextField(hint_text='Umsatzzuteilung', write_tab=False)
            self.create_service_grid_layout.add_widget(self.service_type)
            self.create_service_grid_layout.add_widget(self.description)
            self.create_service_grid_layout.add_widget(self.consulter)
            self.create_service_grid_layout.add_widget(self.customer)
            self.create_service_grid_layout.add_widget(self.abrechnugszyklus)
            self.create_service_grid_layout.add_widget(self.abrechnungstyp)

            # self.create_service_grid_layout.add_widget(self.responsible)
            self.service_type.bind(focus=self.get_service_types)
            self.consulter.bind(focus=self.get_consulter)
            self.customer.bind(focus=self.get_customer)
            self.abrechnugszyklus.bind(focus=self.get_abrechnungszyklusliste)
            self.abrechnungstyp.bind(focus=self.get_abrechnungstypliste)
            # self.responsible.bind(focus=self.get_responsible)

            # Bestätigungs-Button erstellen
            confirm_button = MDRectangleFlatButton(text='Bestätigen', size_hint=(0.2, None))
            confirm_button.bind(on_release=self.save_service)
            self.create_service_button_grid.add_widget(confirm_button)

            # Abbrechen-Button erstellen
            cancel_button = MDRectangleFlatButton(text='Abbrechen', size_hint=(0.2, None))
            cancel_button.bind(on_release=self.cancel_service)
            self.create_service_button_grid.add_widget(cancel_button)

            # GridLayout im Hauptbildschirm anzeigen
            self.add_widget(self.create_service_grid_layout)
            self.add_widget(self.create_service_button_grid)

    def save_service(self, *args):
        if self.service_type.text and self.description.text and self.consulter.text and self.customer.text:
            # neu dazu gekommen, es muss der Fachbereich ermittelt werden
            # dazu wird service_type in dropdown_first_level gesucht
            areas = self.sql_statements.get_servicetypes()
            my_area = ""
            for area in areas:  # Fachbereich
                if area[0] == self.service_type.text:
                    my_area = area[5]
                    break
            self.sql_statements.input_service(self.service_type.text, self.description.text, self.consulter.text,
                                              self.customer.text, my_area, self.abrechnungstyp.text, self.abrechnugszyklus.text)
            data = self.sql_statements.get_services()
            self.service_table.row_data = data
            self.remove_widget(self.create_service_grid_layout)
            self.remove_widget(self.create_service_button_grid)
            self.service_selected = False
            self.create_service_dialog = False
            self.service_view_label.text = "..."
        else:
            self.service_error_label.text = "Es müssen alle Felder gefüllt werden"
            Clock.schedule_once(self.clear_service_error_label, 2)

    def update_service(self, *args):
        if self.service_type.text and self.description.text and self.consulter.text and self.customer.text:
            self.sql_statements.update_service(self.service_type.text, self.description.text, self.consulter.text,
                                               self.customer.text, self.service_info, self.abrechnungstyp.text, self.abrechnugszyklus.text)
            data = self.sql_statements.get_services()
            self.service_table.row_data = data
            self.remove_widget(self.create_service_grid_layout)
            self.remove_widget(self.create_service_button_grid)
            self.service_selected = False
            self.create_service_dialog = False
            self.service_view_label.text = "..."

        else:
            self.service_error_label.text = "Es müssen alle Felder gefüllt werden"
            Clock.schedule_once(self.clear_service_error_label, 2)

    def service_type_multi_filter(self, instance, *args):
        result = self.sql_statements.get_servicetypes_from_db()
        self.list = [
            {
                "viewclass": "OneLineListItem",
                "text": f"{i[0]}",
                "on_release": lambda x=f"{i[0]}": self.service_filter_field(x)
            }
            for i in result
        ]
        # Dropdown wird technisch erzeugt
        self.dropdown = MDDropdownMenu(
            items=self.list,
            width_mult=4,
            position="auto"
        )
        self.dropdown.caller = instance
        self.dropdown.open()

    def service_filter_field(self, instance, *args):
        self.service_type_filter.text = instance
        self.dropdown.dismiss()

    def consulter_multi_filter(self, instance, *args):
        result = self.sql_statements.get_consulter_from_db()
        self.list = [
            {"viewclass": "OneLineListItem",
             "text": f"{i[0]}",
             "on_release": lambda x=f"{i[0]}": self.fill_service_filter_consulter(x)}
            for i in result]
        # DropDown wird technisch erzeugt
        self.dropdown = MDDropdownMenu(
            items=self.list,
            width_mult=4,
            position="auto"
        )
        self.dropdown.caller = instance
        self.dropdown.open()

    def fill_service_filter_consulter(self, instance, *args):
        self.service_consulter_filter.text = instance
        self.dropdown.dismiss()

    def customer_multi_filter(self, instance, *args):
        result = self.sql_statements.get_customer_from_db()
        self.list = [
            {"viewclass": "OneLineListItem",
             "text": f"{i[0]}",
             "on_release": lambda x=f"{i[0]}": self.fill_service_filter_customer(x)}
            for i in result]
        # DropDown wird technisch erzeugt
        self.dropdown = MDDropdownMenu(
            items=self.list,
            width_mult=4,
            position="auto"
        )
        self.dropdown.caller = instance
        self.dropdown.open()

    def fill_service_filter_customer(self, instance, *args):
        self.service_customer_filter.text = instance
        self.dropdown.dismiss()

    def status_multi_filter(self, instance, *args):
        result = self.sql_statements.get_status_from_db()
        self.list = [
            {"viewclass": "OneLineListItem",
             "text": f"{i[0]}",
             "on_release": lambda x=f"{i[0]}": self.fill_service_filter_status(x)}
            for i in result]
        # DropDown wird technisch erzeugt
        self.dropdown = MDDropdownMenu(
            items=self.list,
            width_mult=4,
            position="auto"
        )
        self.dropdown.caller = instance
        self.dropdown.open()

    def fill_service_filter_status(self, instance, *args):
        self.service_status_filter.text = instance
        self.dropdown.dismiss()

    def use_multifilter(self, instance, *args):
        attributes = {}
        attributes = {
            "type": self.service_type_filter.text,
            "consulter": self.service_consulter_filter.text,
            "customer_name": self.service_customer_filter.text,
            "status_type_a": self.service_status_filter.text,
        }
        result = self.sql_statements.filter_datatable(attributes)
        self.service_table.row_data = result

    def reset_filter(self, instance, *args):
        result = self.sql_statements.get_services()
        self.service_table.row_data = result
        self.service_type_filter.text = "..."
        self.service_consulter_filter.text = "..."
        self.service_customer_filter.text = "..."
        self.service_status_filter.text = "..."
        self.informations.text = "Es wurde alle Filter gelöscht"
        Clock.schedule_once(self.clear_information, 2)

    def clear_information(self, instance):
        self.informations.text = "..."

    def clear_service_error_label(self, *args):
        self.service_error_label.text = 'Dienstleistung anlegen'

    def clear_position_error_label(self, *args):
        self.position_error_label.text = 'Position anlegen'

    def get_service_types(self, instance, *args):
        if self.service_type.focus:
            result = self.sql_statements.get_servicetypes()
            self.list = [
                {"viewclass": "OneLineListItem",
                 "text": f"{i[0]}",
                 "on_release": lambda x=f"{i[0]}": self.fill_servicetype_field(x)}
                for i in result]
            # DropDown wird technisch erzeugt
            self.dropdown = MDDropdownMenu(
                items=self.list,
                width_mult=4,
                position="auto"
            )
            self.dropdown.caller = instance
            self.dropdown.open()

    def fill_servicetype_field(self, instance, *args):
        self.dropdown.dismiss()
        self.service_type.text = instance

    def get_consulter(self, instance, *args):
        if self.consulter.focus:
            result, text = self.sql_statements.get_consulter(instance.text)
            self.list = [
                {"viewclass": "OneLineListItem",
                 "text": f"{i[0]}",
                 "on_release": lambda x=f"{i[0]}": self.fill_consulter_field(x)}
                for i in result]
            # DropDown wird technisch erzeugt
            self.dropdown = MDDropdownMenu(
                items=self.list,
                width_mult=4,
                position="auto"
            )
            self.dropdown.caller = instance
            self.dropdown.open()

    def fill_consulter_field(self, instance, *args):
        self.dropdown.dismiss()
        self.consulter.text = instance

    def get_customer(self, instance, *args):
        if self.customer.focus:
            result, text = self.sql_statements.get_customer(instance.text)
            self.list = [
                {"viewclass": "OneLineListItem",
                 "text": f"{i[1]}",
                 "on_release": lambda x=f"{i[1]}": self.fill_customer_field(x)}
                for i in result]
            # DropDown wird technisch erzeugt
            self.dropdown = MDDropdownMenu(
                items=self.list,
                width_mult=4,
                position="auto"
            )
            self.dropdown.caller = instance
            self.dropdown.open()

    def fill_customer_field(self, instance, *args):
        self.dropdown.dismiss()
        self.customer.text = instance

    def get_abrechnungszyklusliste(self, instance, *args):
        if self.abrechnugszyklus.focus:
            result = [
                "jährlich",
                "halbjahrig",
                "quartalsweise",
                "monatlich",
            ]
            self.list = [
                {"viewclass": "OneLineListItem",
                 "text": f"{i}",
                 "on_release": lambda x=f"{i}": self.fill_zyklusfeld(x)}
                for i in result]
            # DropDown wird technisch erzeugt
            self.dropdown = MDDropdownMenu(
                items=self.list,
                width_mult=4,
                position="auto"
            )
            self.dropdown.caller = instance
            self.dropdown.open()

    def fill_zyklusfeld(self, instance, *args):
        self.dropdown.dismiss()
        self.abrechnugszyklus.text = instance

    def get_abrechnungstypliste(self, instance, *args):
        if self.abrechnungstyp.focus:
            result = [
                "SEPA-Lastschrift",
                "manuelle Überweisung",
            ]
            self.list = [
                {"viewclass": "OneLineListItem",
                 "text": f"{i}",
                 "on_release": lambda x=f"{i}": self.fill_abrechnungstyp(x)}
                for i in result]
            # DropDown wird technisch erzeugt
            self.dropdown = MDDropdownMenu(
                items=self.list,
                width_mult=4,
                position="auto"
            )
            self.dropdown.caller = instance
            self.dropdown.open()

    def fill_abrechnungstyp(self, instance, *args):
        self.dropdown.dismiss()
        self.abrechnungstyp.text = instance

    def get_responsible(self, instance, *args):
        if self.responsible.focus:
            result, text = self.sql_statements.get_responsible(instance.text)
            self.list = [
                {"viewclass": "OneLineListItem",
                 "text": f"{i[1]}",
                 "on_release": lambda x=f"{i[1]}": self.fill_responsible_field(x)}
                for i in result]
            # DropDown wird technisch erzeugt
            self.dropdown = MDDropdownMenu(
                items=self.list,
                width_mult=4,
                position="auto"
            )
            self.dropdown.caller = instance
            self.dropdown.open()

    def fill_responsible_field(self, instance, *args):
        self.dropdown.dismiss()
        self.responsible.text = instance

    def cancel_service(self, instance):
        self.create_service_dialog = False
        self.remove_widget(self.create_service_grid_layout)
        self.remove_widget(self.create_service_button_grid)

    def edit_service(self, instance):
        if not self.create_service_dialog:  # Überprüfen, ob der Dialog bereits geöffnet ist
            if self.service_selected is True:
                if self.service_info[7] != "abgerechnet" and self.service_info[7] != "abgeschlossen" and \
                        self.service_info[7] != "bezahlt":
                    self.create_service_dialog = True

                    # GridLayout erstellen
                    self.create_service_grid_layout = MDGridLayout(cols=1, padding=10, size_hint=(0.2, 0.3),
                                                                   pos_hint={'x': 0.75, 'y': 0.6})
                    self.create_service_grid_layout.line_width = 1

                    self.create_service_button_grid = MDGridLayout(cols=2, padding=10, size_hint=(0.2, 0.05),
                                                                   pos_hint={'x': 0.75, 'y': 0.5})
                    self.create_service_button_grid.line_width = 1

                    # Fehlermeldungstext erstellen
                    self.service_error_label = MDLabel(text='Dienstleistung ändern')
                    self.create_service_grid_layout.add_widget(self.service_error_label)

                    # Texteingabefelder erstellen
                    self.service_type = MDTextField(hint_text='Dienstleistungstyp', write_tab=False)
                    self.description = MDTextField(hint_text='Bezeichnung', write_tab=False)
                    self.consulter = MDTextField(hint_text='Ansprechpartner', write_tab=False)
                    self.customer = MDTextField(hint_text='Rechnungsempfänger', write_tab=False)
                    self.abrechnugszyklus = MDTextField(hint_text='Abrechnungszyklus', write_tab=False)
                    self.abrechnungstyp = MDTextField(hint_text='Abrechnungstyp', write_tab=False)
                    # self.responsible = MDTextField(hint_text='Umsatzzuteilung', write_tab=False)
                    self.create_service_grid_layout.add_widget(self.service_type)
                    self.create_service_grid_layout.add_widget(self.description)
                    self.create_service_grid_layout.add_widget(self.consulter)
                    self.create_service_grid_layout.add_widget(self.customer)
                    self.create_service_grid_layout.add_widget(self.abrechnugszyklus)
                    self.create_service_grid_layout.add_widget(self.abrechnungstyp)
                    # self.create_service_grid_layout.add_widget(self.responsible)
                    self.service_type.bind(focus=self.get_service_types)
                    self.consulter.bind(focus=self.get_consulter)
                    self.customer.bind(focus=self.get_customer)
                    self.abrechnugszyklus.bind(focus=self.get_abrechnungszyklusliste)
                    self.abrechnungstyp.bind(focus=self.get_abrechnungstypliste)
                    # self.responsible.bind(focus=self.get_responsible)

                    # Bestätigungs-Button erstellen
                    confirm_button = MDRectangleFlatButton(text='Bestätigen', size_hint=(0.2, None))
                    confirm_button.bind(on_release=self.update_service)
                    self.create_service_button_grid.add_widget(confirm_button)

                    # Abbrechen-Button erstellen
                    cancel_button = MDRectangleFlatButton(text='Abbrechen', size_hint=(0.2, None))
                    cancel_button.bind(on_release=self.cancel_service)
                    self.create_service_button_grid.add_widget(cancel_button)

                    self.add_widget(self.create_service_grid_layout)
                    self.add_widget(self.create_service_button_grid)

                    self.service_type.text = self.service_info[1]
                    self.description.text = self.service_info[2]
                    self.consulter.text = self.service_info[5]
                    self.customer.text = self.service_info[6]
                    servicedata = self.sql_statements.get_services_with_id(self.service_info)
                    self.abrechnungstyp.text = servicedata[0][8]
                    self.abrechnugszyklus.text = servicedata[0][9]
                    # self.responsible.text = self.service_info[9]
                else:
                    self.informations.text = "Dienstleistung kann nicht mehr geändert werden"
                    Clock.schedule_once(self.clear_information, 2)
            else:
                self.informations.text = "Es muss ein DLN geklickt werden"
                Clock.schedule_once(self.clear_information, 2)

    def delete_service(self, instance):
        self.informations.text = " Das Löschen einer DLN ist aktuell nicht erlaubt"
        Clock.schedule_once(self.clear_information, 2)

    def view_service(self, instance):
        if self.service_selected is True:
            pdf_data = self.sql_statements.get_services_with_id(self.service_info)
            pdf_data.append(self.sql_statements.get_adress(self.service_info))
            # jetzt werden noch alle Position-Typen zu dieser DLN benötigt
            positiongroups = self.sql_statements.get_positiongroups(self.service_info)
            # mithilfe der positionsgroups und der service id können alle positions gezogen werden
            position_dict = {}
            for position in positiongroups:
                position_items = self.sql_statements.get_positions_with_group_and_id(self.service_info, position)

                # Sortieren der 'position_items' basierend auf dem Datum, das im dritten Element des Tuples steht
                position_items.sort(key=lambda item: datetime.strptime(item[2], '%d.%m.%Y'))

                position_dict[position] = position_items
            customer = pdf_data[0][4]
            try:
                customer_adress = pdf_data[1][0][28]
                customer_code = pdf_data[1][0][30]
                customer_town = pdf_data[1][0][29]
            except:
                customer_adress = ""
                customer_code = ""
                customer_town = ""
            service_type = pdf_data[0][6]
            description = pdf_data[0][1]
            service_start = pdf_data[0][2]
            service_end = pdf_data[0][3]
            service_number = pdf_data[0][0]

            generate_pdf = MyDocTemplate()
            generate_pdf.create_pdf(position_dict, customer, customer_adress, customer_code, customer_town,
                                    service_type, description, service_start, service_end, service_number)
            self.service_view_label.text = ""
            self.service_selected = False

    def send_service(self, instance):
        # Aktion für "Dienstleistung senden" ausführen
        pass

    def change_status(self, instance):
        if self.service_selected:
            if self.role == 'admin' or self.role == 'backoffice':
                self.list = [
                    {"viewclass": "OneLineListItem",
                     "text": "bearbeitet",
                     "on_release": lambda x="bearbeitet": self.handle_status(x)},

                    {"viewclass": "OneLineListItem",
                     "text": "abgeschlossen",
                     "on_release": lambda x="abgeschlossen": self.handle_status(x)},

                    {"viewclass": "OneLineListItem",
                     "text": "abgerechnet",
                     "on_release": lambda x="abgerechnet": self.handle_status(x)},
                ]
                self.dropdown = MDDropdownMenu(
                    items=self.list,
                    width_mult=4
                )
                self.dropdown.caller = instance
                self.dropdown.open()

            if self.role == 'beratung' or self.role == 'bewertung':
                self.list = [
                    {"viewclass": "OneLineListItem",
                     "text": "bearbeitet",
                     "on_release": lambda x="bearbeitet": self.handle_status(x)},

                    {"viewclass": "OneLineListItem",
                     "text": "abgeschlossen",
                     "on_release": lambda x="abgeschlossen": self.handle_status(x)},
                ]
                self.dropdown = MDDropdownMenu(
                    items=self.list,
                    width_mult=4
                )
                self.dropdown.caller = instance
                self.dropdown.open()
        else:
            self.informations.text = "Es muss ein DLN ausgewählt werden"
            Clock.schedule_once(self.clear_information, 2)

    def handle_status(self, status):
        self.sql_statements.change_status(status, self.service_info)
        servicedata = self.sql_statements.get_services()
        self.service_table.row_data = servicedata
        self.dropdown.dismiss()

    def create_position(self, instance):
        if self.create_position_dialog is False:
            if self.service_selected is True:
                if self.service_info[7] != "abgerechnet" and self.service_info[7] != "abgeschlossen" and \
                        self.service_info[7] != "bezahlt":
                    self.create_position_dialog = True
                    # GridLayout erstellen
                    self.create_position_grid_layout = MDGridLayout(cols=2, padding=2, spacing=20, size_hint=(0.2, 0.3),
                                                                    pos_hint={'x': 0.75, 'y': 0.15})

                    self.create_position_button_grid = MDGridLayout(cols=2, padding=6, size_hint=(0.2, 0.05),
                                                                    pos_hint={'x': 0.75, 'y': 0.05})

                    # Fehlermeldungstext erstellen
                    self.position_error_label = MDLabel(text='Position anlegen')
                    # self.create_position_grid_layout.add_widget(self.position_error_label)

                    # Texteingabefelder erstellen
                    self.position_type = MDTextField(hint_text='Dienstleistungstyp', write_tab=False, )
                    self.position_description = MDTextField(hint_text='Bezeichnung', write_tab=False)
                    self.position_comment = MDTextField(hint_text='Bemerkung', write_tab=False)
                    self.position_datefield = MDTextField(hint_text='Datum der Leistung', write_tab=False)
                    self.position_amount = MDTextField(hint_text='Menge', write_tab=False)
                    self.position_singleprice = MDTextField(hint_text='Einzelpreis', write_tab=False)
                    self.position_consulter = MDTextField(hint_text='Ansprechpartner', write_tab=False)
                    self.responsible = MDTextField(hint_text='Umsatzzuteilung', write_tab=False, )
                    self.create_position_grid_layout.add_widget(self.position_type)
                    self.create_position_grid_layout.add_widget(self.position_description)
                    self.create_position_grid_layout.add_widget(self.position_comment)
                    self.create_position_grid_layout.add_widget(self.position_datefield)
                    self.create_position_grid_layout.add_widget(self.position_amount)
                    self.create_position_grid_layout.add_widget(self.position_singleprice)
                    self.create_position_grid_layout.add_widget(self.position_consulter)
                    self.create_position_grid_layout.add_widget(self.responsible)

                    self.position_type.bind(focus=self.get_position_types)
                    self.position_description.bind(focus=self.get_position_description)
                    self.position_datefield.bind(focus=self.get_position_date)
                    self.position_consulter.bind(focus=self.get_position_consulter)
                    self.responsible.bind(focus=self.get_responsible)

                    # Bestätigungs-Button erstellen
                    confirm_button = MDRectangleFlatButton(text='Bestätigen', size_hint=(0.2, None))
                    confirm_button.bind(on_release=self.save_position)
                    self.create_position_button_grid.add_widget(confirm_button)

                    # Abbrechen-Button erstellen
                    cancel_button = MDRectangleFlatButton(text='Abbrechen', size_hint=(0.2, None))
                    cancel_button.bind(on_release=self.cancel_position)
                    self.create_position_button_grid.add_widget(cancel_button)

                    # GridLayout im Hauptbildschirm anzeigen
                    self.add_widget(self.create_position_grid_layout)
                    self.add_widget(self.create_position_button_grid)
                else:
                    self.informations.text = "Eine Bearbeitung dieser DLN ist nicht mehr möglich"
                    Clock.schedule_once(self.clear_information, 2)
            else:
                self.informations.text = "Es muss ein DLN ausgewählt werden"
                Clock.schedule_once(self.clear_information, 2)

    def edit_positions(self, instance):
        if self.create_position_dialog is not True:
            if self.position_selected is True:
                if self.service_info[7] != "abgerechnet" and self.service_info[7] != "abgeschlossen" and \
                        self.service_info[7] != "bezahlt":
                    self.create_position_dialog = True
                    # GridLayout erstellen
                    self.create_position_grid_layout = MDGridLayout(cols=2, padding=2, spacing=20, size_hint=(0.2, 0.3),
                                                                    pos_hint={'x': 0.75, 'y': 0.15})

                    self.create_position_button_grid = MDGridLayout(cols=2, padding=6, size_hint=(0.2, 0.05),
                                                                    pos_hint={'x': 0.75, 'y': 0.05})

                    # Fehlermeldungstext erstellen
                    self.position_error_label = MDLabel(text='Position bearbeiten')
                    # self.create_position_grid_layout.add_widget(self.position_error_label)

                    # Texteingabefelder erstellen
                    self.position_type = MDTextField(hint_text='Dienstleistungstyp', write_tab=False)
                    self.position_description = MDTextField(hint_text='Bezeichnung', write_tab=False)
                    self.position_comment = MDTextField(hint_text='Bemerkung', write_tab=False)
                    self.position_datefield = MDTextField(hint_text='Datum der Leistung', write_tab=False)
                    self.position_amount = MDTextField(hint_text='Menge', write_tab=False)
                    self.position_singleprice = MDTextField(hint_text='Einzelpreis', write_tab=False)
                    self.position_consulter = MDTextField(hint_text='Ansprechpartner', write_tab=False)
                    self.responsible = MDTextField(hint_text='Umsatzzuteilung', write_tab=False)
                    self.create_position_grid_layout.add_widget(self.position_type)
                    self.create_position_grid_layout.add_widget(self.position_description)
                    self.create_position_grid_layout.add_widget(self.position_comment)
                    self.create_position_grid_layout.add_widget(self.position_datefield)
                    self.create_position_grid_layout.add_widget(self.position_amount)
                    self.create_position_grid_layout.add_widget(self.position_singleprice)
                    self.create_position_grid_layout.add_widget(self.position_consulter)
                    self.create_position_grid_layout.add_widget(self.responsible)

                    self.position_type.bind(focus=self.get_position_types)
                    self.position_description.bind(focus=self.get_position_description)
                    self.position_datefield.bind(focus=self.get_position_date)
                    self.position_consulter.bind(focus=self.get_position_consulter)
                    self.responsible.bind(focus=self.get_responsible)

                    # Bestätigungs-Button erstellen
                    confirm_button = MDRectangleFlatButton(text='Bestätigen', size_hint=(0.2, None))
                    confirm_button.bind(on_release=self.update_positions)
                    self.create_position_button_grid.add_widget(confirm_button)

                    # Abbrechen-Button erstellen
                    cancel_button = MDRectangleFlatButton(text='Abbrechen', size_hint=(0.2, None))
                    cancel_button.bind(on_release=self.cancel_position)
                    self.create_position_button_grid.add_widget(cancel_button)

                    # GridLayout im Hauptbildschirm anzeigen
                    self.add_widget(self.create_position_grid_layout)
                    self.add_widget(self.create_position_button_grid)

                    self.position_type.text = self.position_info[1]
                    self.position_description.text = self.position_info[2]
                    self.position_comment.text = self.position_info[3]
                    self.position_datefield.text = self.position_info[4]
                    self.position_amount.text = str(self.position_info[5])
                    self.position_singleprice.text = str(self.position_info[6])
                    self.position_consulter.text = self.position_info[8]
                    try:
                        self.responsible.text = self.position_info[9]
                    except:
                        self.responsible.text = ""
                else:
                    self.informations.text = "Eine Bearbeitung dieser DLN ist nicht mehr möglich"
                    Clock.schedule_once(self.clear_information, 2)
            else:
                self.informations.text = "Es muss ein Position ausgewählt werden"
                Clock.schedule_once(self.clear_information, 2)

    def save_position(self, *args):
        if self.position_type.text and self.position_description.text and self.position_datefield.text and self.position_amount.text and self.position_singleprice.text and self.position_consulter.text:
            self.sql_statements.input_and_calculate_positions(self.position_type.text, self.position_description.text,
                                                              self.position_datefield.text, self.position_amount.text,
                                                              self.position_singleprice.text,
                                                              self.position_consulter.text, self.service_info,
                                                              self.position_comment.text, self.responsible.text)

            self.remove_widget(self.create_position_grid_layout)
            self.remove_widget(self.create_position_button_grid)
            self.create_position_dialog = False
            status = "bearbeitet"
            self.sql_statements.change_status(status, self.service_info)

            self.calculate_service_dates()
        else:
            self.position_error_label.text = "Es müssen alle Felder gefüllt werden"
            Clock.schedule_once(self.clear_position_error_label, 2)

    def calculate_service_dates(self):
        # es müssen alle Positionen zu diesem Service geladen werden
        positions = self.sql_statements.get_positions(self.service_info)
        position_date_list = [x[4] for x in positions]
        dates = [datetime.strptime(date_string, "%d.%m.%Y") for date_string in
                 position_date_list]  # Dies ersetzt die Schleife und fügt alle Daten hinzu

        # Überprüfen, ob die Liste nicht leer ist, bevor min und max berechnet werden
        if position_date_list:
            service_min_date = min(dates)
            service_max_date = max(dates)
        try:
            self.sql_statements.update_service_header(self.service_info, service_min_date, service_max_date)
        except:
            print("Input Datumswerte fehlgeschlagen")

        service_data = self.sql_statements.get_services()
        self.service_table.row_data = service_data

        position_data = self.sql_statements.get_positions(self.service_info)
        self.position_table.row_data = position_data

    def update_positions(self, *args):
        if self.position_type.text and self.position_description.text and self.position_datefield.text and self.position_amount.text and self.position_singleprice.text and self.position_consulter.text:

            self.sql_statements.update_positions(self.position_type.text, self.position_description.text,
                                                 self.position_datefield.text, self.position_amount.text,
                                                 self.position_singleprice.text, self.position_consulter.text,
                                                 self.position_info, self.position_comment.text, self.responsible.text)

            data = self.sql_statements.get_positions(self.service_info)
            self.position_table.row_data = data
            self.remove_widget(self.create_position_grid_layout)
            self.remove_widget(self.create_position_button_grid)
            self.create_position_dialog = False
            self.position_selected = False
            self.position_view_label.text = ""
            self.calculate_service_dates()
        else:
            self.service_error_label.text = "Es müssen alle Felder gefüllt werden"
            Clock.schedule_once(self.clear_service_error_label, 2)

    def get_position_types(self, instance, *args):
        # es wurde eine Try Schleife eingebaut, damit das Programm nicht abstürzt, wenn der Servicetyp nicht im
        # Dropdown zu finden ist
        try:
            if self.position_type.focus:
                result = self.sql_statements.get_positiontypes(self.service_info)
                self.list = [
                    {"viewclass": "OneLineListItem",
                     "text": f"{i[0]}",
                     "on_release": lambda x=f"{i[0]}": self.fill_positiontype_field(x)}
                    for i in result]
                # DropDown wird technisch erzeugt
                self.dropdown = MDDropdownMenu(
                    items=self.list,
                    width_mult=4,
                    position="auto"
                )
                self.dropdown.caller = instance
                self.dropdown.open()
        except:
            self.position_type.text = ""

    def fill_positiontype_field(self, instance, *args):
        self.dropdown.dismiss()
        self.position_type.text = instance

    def get_position_description(self, instance, *args):
        if self.position_type.text != "":
            try:
                print("Starte SQL Statement")
                result, value = self.sql_statements.get_positiondescription(self.position_type.text)
                self.list = [
                    {"viewclass": "OneLineListItem",
                     "text": f"{i[0]}",
                     "on_release": lambda x=f"{i[0]}": self.fill_positiondescription_field(x, value, result)}
                    for i in result]
                # DropDown wird technisch erzeugt
                self.dropdown = MDDropdownMenu(
                    items=self.list,
                    width_mult=4,
                    position="auto"
                )
                self.dropdown.caller = instance
                self.dropdown.open()
            except:
                self.position_description.text = ""
                self.position_singleprice.text = ""
                print("Try Schleife")
        else:
            print("Bitte oben anfangen")

    def fill_positiondescription_field(self, instance, value, result, *args):
        self.dropdown.dismiss()
        self.position_description.text = instance
        self.position_singleprice.text = value
        # Finde den Index des Tuples mit dem Wert 'instance'
        index = next((i for i, t in enumerate(result) if t[0] == instance), None)

        # Wenn das Tuple gefunden wurde und die zweite Stelle (Index 1) des Tuples gleich '.' ist
        if index is not None and result[index][1] == '.':
            self.responsible.text = "admin"
        else:
            self.responsible.text = self.service_info[5]

    def get_position_date(self, instance, *args):
        if self.position_datefield.focus:
            self.date_dialog = MDDatePicker()
            self.date_dialog.bind(on_save=self.fill_date_field)
            self.date_dialog.open()

    def fill_date_field(self, instance, value, date_range):
        self.position_datefield.text = str(value)

    def get_position_consulter(self, instance, *args):
        if self.position_consulter.focus:
            result, text = self.sql_statements.get_consulter(instance.text)
            self.list = [
                {"viewclass": "OneLineListItem",
                 "text": f"{i[0]}",
                 "on_release": lambda x=f"{i[0]}": self.fill_positionconsulter_field(x)}
                for i in result]
            # DropDown wird technisch erzeugt
            self.dropdown = MDDropdownMenu(
                items=self.list,
                width_mult=4,
                position="auto"
            )
            self.dropdown.caller = instance
            self.dropdown.open()

    def fill_positionconsulter_field(self, instance, *args):
        self.dropdown.dismiss()
        self.position_consulter.text = instance

    def cancel_position(self, instance):
        self.create_position_dialog = False
        try:
            self.remove_widget(self.create_position_grid_layout)
        except:
            pass
        try:
            self.remove_widget(self.create_position_button_grid)
        except:
            pass

    def edit_position(self, instance):
        pass

    def delete_position(self, instance):
        if self.position_selected and self.create_position_dialog == False:
            self.create_position_dialog = True
            self.create_position_button_grid = MDGridLayout(cols=2, padding=6, size_hint=(0.2, 0.05),
                                                            pos_hint={'x': 0.75, 'y': 0.1},
                                                            line_color="black")

            # Bestätigungs-Button erstellen
            confirm_button = MDRectangleFlatButton(text='Löschen', size_hint=(0.2, None))
            confirm_button.bind(on_release=self.set_delete_status)
            self.create_position_button_grid.add_widget(confirm_button)

            # Abbrechen-Button erstellen
            cancel_button = MDRectangleFlatButton(text='Abbrechen', size_hint=(0.2, None))
            cancel_button.bind(on_release=self.cancel_position)
            self.create_position_button_grid.add_widget(cancel_button)

            self.add_widget(self.create_position_button_grid)
        else:
            self.informations.text = "Es muss erst eine Position ausgewählt werden"
            Clock.schedule_once(self.clear_information, 2)

    def set_delete_status(self, instance):
        status = "gelöscht"
        self.sql_statements.delete_position(self.position_info)
        self.calculate_service_dates()
        self.remove_widget(self.create_position_button_grid)
        self.position_view_label.text = "..."
        self.position_selected = False

    def dropdown_rooms(self, instance):
        self.list = [
            {"viewclass": "OneLineListItem",
             "text": "Dienstleistungen",
             "on_release": lambda x="services": self.activate_room(x)},

            {"viewclass": "OneLineListItem",
             "text": "Adminbereich",
             "on_release": lambda x="admin": self.activate_room(x)},

        ]
        self.dropdown = MDDropdownMenu(
            items=self.list,
            width_mult=4
        )
        self.dropdown.caller = instance
        self.dropdown.open()

    def activate_room(self, room="Services"):
        if self.role == 'admin' or self.role == 'backoffice':
            agrar.screen_manager.current = room
            self.dropdown.dismiss()


class AdminScreen(Screen):
    def __init__(self, **kwargs):
        super(AdminScreen, self).__init__(**kwargs)
        self.general_view_label = None
        self.dunning_row_selected = None
        self.dunning_table_created = False
        self.name = "admin"
        self.sqlstatements = SQLStatements(mandant)

        toolbar = MDTopAppBar(
            pos_hint={'top': 1},
            title="",
            md_bg_color=[0.18, 0.19, 0.19, 1],
            left_action_items=[["dots-vertical", lambda x: self.dropdown_rooms(x)],
                               ["account-arrow-up-outline", lambda x: self.import_crm_data(x),
                                "CRM-Daten importieren [ca. 5 Minuten]"],
                               ["table-large-plus", lambda x: self.create_dunning_table(x), "Mahnungstabelle öffnen"],
                               ["table-large-remove", lambda x: self.close_dunning_table(x),
                                "Mahnungstabelle schließen"],
                               ["", lambda x: "", "", ""],
                               ["table-key", lambda x: self.id_dunning_filter(x), "Nummernfilter"],  # fertig
                               ["table-search", lambda x: self.man_dunning_multi_filter(x),  # fertig
                                "Filter für gesetzte Mahnstufen"],
                               ["table-search", lambda x: self.auto_dunning_multi_filter(x),
                                "Filter für automatische Mahnstufen"],  # fertig
                               ["account-question", lambda x: self.consulter_dunning_multi_filter(x),
                                "Ansprechpartner auswählen"],
                               ["account-question-outline", lambda x: self.customer_dunning_multi_filter(x),
                                "Rechnungsempfänger auswählen"],
                               ["filter-check-outline", lambda x: self.use_multifilter(x), "Filter anwenden"],
                               ["filter-remove-outline", lambda x: self.reset_filter(x), "Filter zurücksetzen"],

                               ],
            right_action_items=[["table-edit", lambda x: self.change_dunning_status(x), "Mahnungsstatus ändern"],
                                ["table-refresh", lambda x: self.calculate_dunning_stages(x),
                                 "Mahnstufen ausrechnen [einige Sekunden]"],
                                ["table-check", lambda x: self.set_account_date(x), "DLN auf bezahlt setzen"],
                                ]
        )

        self.add_widget(toolbar)

    def id_dunning_filter(self, instance, *args):
        # dropdown aller Nummer der Services
        if self.dunning_table_created:
            result = self.sqlstatements.get_services_for_dunning()
            # Liste nach dem ersten Eintrag jeder Position sortieren
            sorted_list = sorted(result, key=lambda x: x[0], reverse=True)
            self.list = [
                {
                    "viewclass": "OneLineListItem",
                    "text": f"{i[0]}",
                    "on_release": lambda x=f"{i[0]}": self.filter_servicetable_with_number(x, sorted_list)
                }
                for i in sorted_list
            ]
            # Dropdown wird technisch erzeugt
            self.dropdown = MDDropdownMenu(
                items=self.list,
                width_mult=4,
                position="auto"
            )
            self.dropdown.caller = instance
            self.dropdown.open()
        else:
            self.general_view_label.text = "Bitte erst die Tabelle öffnen"
            Clock.schedule_once(self.clear_general_view_label, 3)

    def filter_servicetable_with_number(self, instance, sorted_list):
        data = [eintrag for eintrag in sorted_list if eintrag[0] == int(instance)]
        print(data)
        self.dunning_table.row_data = data
        self.dropdown.dismiss()
        self.general_view_label.text = "..."
        self.dunning_row_selected = False

    def man_dunning_multi_filter(self, instance, *args):
        if self.dunning_table_created:
            result = self.sqlstatements.get_servicedunning_man_from_db()
            self.list = [
                {
                    "viewclass": "OneLineListItem",
                    "text": f"{i[0]}",
                    "on_release": lambda x=f"{i[0]}": self.fill_dunning_manuel(x)
                }
                for i in result
            ]
            # Dropdown wird technisch erzeugt
            self.dropdown = MDDropdownMenu(
                items=self.list,
                width_mult=4,
                position="auto"
            )
            self.dropdown.caller = instance
            self.dropdown.open()
        else:
            self.general_view_label.text = "Bitte erst die Tabelle öffnen"
            Clock.schedule_once(self.clear_general_view_label, 3)

    def fill_dunning_manuel(self, instance, *args):
        self.man_dunning.text = instance
        self.dropdown.dismiss()

    def auto_dunning_multi_filter(self, instance, *args):
        if self.dunning_table_created:
            result = self.sqlstatements.get_servicedunning_auto_from_db()
            self.list = [
                {
                    "viewclass": "OneLineListItem",
                    "text": f"{i[0]}",
                    "on_release": lambda x=f"{i[0]}": self.fill_dunning_auto(x)
                }
                for i in result
            ]
            # Dropdown wird technisch erzeugt
            self.dropdown = MDDropdownMenu(
                items=self.list,
                width_mult=4,
                position="auto"
            )
            self.dropdown.caller = instance
            self.dropdown.open()
        else:
            self.general_view_label.text = "Bitte erst die Tabelle öffnen"
            Clock.schedule_once(self.clear_general_view_label, 3)

    def fill_dunning_auto(self, instance, *args):
        self.auto_dunning.text = instance
        self.dropdown.dismiss()

    def consulter_dunning_multi_filter(self, instance, *args):
        if self.dunning_table_created:
            result = self.sqlstatements.get_consulter_from_db_1()
            self.list = [
                {
                    "viewclass": "OneLineListItem",
                    "text": f"{i[0]}",
                    "on_release": lambda x=f"{i[0]}": self.fill_consulter(x)
                }
                for i in result
            ]
            # Dropdown wird technisch erzeugt
            self.dropdown = MDDropdownMenu(
                items=self.list,
                width_mult=4,
                position="auto"
            )
            self.dropdown.caller = instance
            self.dropdown.open()
        else:
            self.general_view_label.text = "Bitte erst die Tabelle öffnen"
            Clock.schedule_once(self.clear_general_view_label, 3)

    def fill_consulter(self, instance, *args):
        self.consulter_dunning.text = instance
        self.dropdown.dismiss()

    def customer_dunning_multi_filter(self, instance, *args):
        if self.dunning_table_created:
            result = self.sqlstatements.get_customer_from_db_1()
            self.list = [
                {
                    "viewclass": "OneLineListItem",
                    "text": f"{i[0]}",
                    "on_release": lambda x=f"{i[0]}": self.fill_customer(x)
                }
                for i in result
            ]
            # Dropdown wird technisch erzeugt
            self.dropdown = MDDropdownMenu(
                items=self.list,
                width_mult=4,
                position="auto"
            )
            self.dropdown.caller = instance
            self.dropdown.open()
        else:
            self.general_view_label.text = "Bitte erst die Tabelle öffnen"
            Clock.schedule_once(self.clear_general_view_label, 3)

    def fill_customer(self, instance, *args):
        self.customer_dunning.text = instance
        self.dropdown.dismiss()

    def use_multifilter(self, instance, *args):
        attributes = {}
        attributes = {
            "status_type_c": self.man_dunning.text,
            "consulter": self.consulter_dunning.text,
            "customer_name": self.customer_dunning.text,
            "status_type_b": self.auto_dunning.text,
        }
        result = self.sqlstatements.filter_datatable_1(attributes)
        self.dunning_table.row_data = result

    def reset_filter(self, instance, *args):
        result = self.sqlstatements.get_services_for_dunning()
        self.dunning_table.row_data = result
        self.auto_dunning.text = "..."
        self.man_dunning.text = "..."
        self.customer_dunning.text = "..."
        self.consulter_dunning.text = "..."
        self.general_view_label.text = "Es wurde alle Filter gelöcht"
        Clock.schedule_once(self.clear_general_view_label, 2)

    def dropdown_rooms(self, instance):
        self.list = [
            {"viewclass": "OneLineListItem",
             "text": "Dienstleistungen",
             "on_release": lambda x="services": self.activate_room(x)},

            {"viewclass": "OneLineListItem",
             "text": "Adminbereich",
             "on_release": lambda x="admin": self.activate_room(x)},

        ]
        self.dropdown = MDDropdownMenu(
            items=self.list,
            width_mult=4
        )
        self.dropdown.caller = instance
        self.dropdown.open()

    def on_leave(self):
        self.dunning_row_selected = False
        self.dunning_table_created = False

        # Liste der zu überprüfenden Klassen
        classes_to_remove = (MDDataTable, MDLabel, MDGridLayout)

        # Entfernen der MDDataTable und MDLabel aus dem Widget-Tree
        widgets_to_remove = [child for child in self.children if isinstance(child, classes_to_remove)]
        for widget in widgets_to_remove:
            self.remove_widget(widget)

    def on_enter(self, *args):
        self.filter_grid = MDGridLayout(cols=5, line_color="white", pos_hint={'x': 0.35, 'y': 0.95},
                                        size_hint=(0.5, 0.04), padding=5)

        self.man_dunning = MDLabel(text='...', theme_text_color='Custom', text_color=(1, 1, 1, 1))
        self.auto_dunning = MDLabel(text='...', theme_text_color='Custom', text_color=(1, 1, 1, 1))
        self.consulter_dunning = MDLabel(text='...', theme_text_color='Custom', text_color=(1, 1, 1, 1))
        self.customer_dunning = MDLabel(text='...', theme_text_color='Custom', text_color=(1, 1, 1, 1))

        self.add_widget(self.filter_grid)
        self.filter_grid.add_widget(self.man_dunning)
        self.filter_grid.add_widget(self.auto_dunning)
        self.filter_grid.add_widget(self.consulter_dunning)
        self.filter_grid.add_widget(self.customer_dunning)

        self.general_view_label = MDLabel(text="Hallo", pos_hint={"x": 0.5, "y": 0.42})
        self.add_widget(self.general_view_label)

    def activate_room(self, room="Adminbereich"):
        agrar.screen_manager.current = room
        self.dropdown.dismiss()

    def import_crm_data(self, instance):
        windows_username = os.environ.get('USERNAME')
        self.customer_path = "C:\\Users\\" + windows_username + "\\VR AgrarBeratung AG\\genoBIT Administrator - - VRAgrar Daten\\308 Projekte\\1_HfM\\CRM_Export\\"
        selected_file = plyer.filechooser.open_file(filetypes=[(self.customer_path, '*.xlsx')])
        if selected_file:
            # Extrahiere den ausgewählten Dateipfad
            file_path = selected_file[0]

            # Überprüfen, ob der Dateiname "CentralStationCRM_Company_full" enthält
            if "CentralStationCRM_Company_full" not in os.path.basename(file_path):
                self.general_view_label.text = "Nicht die richtige Datei ausgewählt"
                Clock.schedule_once(self.clear_general_view_label, 5)
                return

            df = pandas.read_excel(file_path)
            df = df.fillna("")
            self.sqlstatements.delete_crm_data()

            for line in df.values:
                customer_id = line[0]
                customer = line[1]
                adress = line[2]
                postalcode = line[3]
                domicilie = line[4]
                phone = line[29]
                verantwortlich = line[49]
                hfm_servernummer = line[60]
                self.sqlstatements.input_customer(customer, customer_id, adress, domicilie, postalcode, phone,
                                                  verantwortlich, hfm_servernummer)

        self.general_view_label.text = "Upload der CRM Daten fertig"
        Clock.schedule_once(self.clear_general_view_label, 5)

    def clear_general_view_label(self, instance):
        self.general_view_label.text = "..."

    def create_dunning_table(self, instance):
        if not self.dunning_table_created:
            #   self.calculate dunnings wird ausgelagert werden
            #self.calculate_dunning_stages(instance)
            table_data = self.sqlstatements.get_services_for_dunning()
            # Beim ersten Aufruf werden die bezahlten rausgefiltert
            filtered_data = self.filter_entries(table_data)

            # Spaltenüberschriften für die Tabelle
            service_column_data = [
                ("DL-ID", Window.width * 0.01),
                ("Abrechnungsdatum", Window.width * 0.017),
                ("Zahldatum", Window.width * 0.015),
                ("Beschreibung", Window.width * 0.035),
                ("Ansprechpartner", Window.width * 0.015),
                ("Rechnungsempfänger", Window.width * 0.025),
                ("Status DLN", Window.width * 0.015),
                ("automatische Mahnstufe", Window.width * 0.02),
                ("gesetzte Mahnstufe", Window.width * 0.02),
            ]
            # MDDataTable erstellen
            self.dunning_table = MDDataTable(
                size_hint=(0.9, 0.4),  # Größe des Tabellenwidgets im Verhältnis zum Screen
                pos_hint={"x": 0.04, "y": 0.5},
                column_data=service_column_data,
                row_data=filtered_data,
                use_pagination=False,  # Bei vielen Daten die Verwendung von Seitenumbrüchen ermöglichen (optional)
                pagination_menu_pos="center",  # Position des Seitenumbruchmenüs (optional)
                rows_num=10000,  # Anzahl der Zeilen pro Seite festlegen
            )
            self.dunning_table.bind(on_row_press=self.on_row_press_dunning)
            self.add_widget(self.dunning_table)
            self.dunning_table_created = True

    def filter_entries(self, table_data):
        data_filtered = []
        for entry in table_data:
            if entry[6] != "bezahlt":
                data_filtered.append(entry)
        return data_filtered

    def close_dunning_table(self, instance):
        try:
            self.remove_widget(self.dunning_table)
            self.general_view_label.text = "Tabelle wurde geschlossen"
            self.dunning_table_created = False
            self.dunning_row_selected = False
        except:
            pass

    def on_row_press_dunning(self, instance_table, instance_row, *args):
        selected_range = instance_row.range
        tabledata = instance_row.table.row_data

        new_data = [value for sublist in tabledata for value in sublist]
        self.dunning_info = []

        for index in range(selected_range[0], selected_range[1] + 1):
            self.dunning_info.append(new_data[index])

        self.general_view_label.text = "Folgender DLN wurde gewählt: " + str(
            self.dunning_info[4]) + " mit der Nr. " + str(self.dunning_info[0])
        self.dunning_row_selected = True

    def calculate_dunning_stages(self, instance):
        # Die ganze Mahnungsdatenbank wird berechnet
        dunning_times = self.sqlstatements.get_dunning_options()
        filtered_data = [item for item in dunning_times if item[3] != '0']

        services = self.sqlstatements.get_services_for_dunning()

        heute = datetime.now()
        datum_string = heute.strftime("%d.%m.%Y")
        today = datetime.strptime(datum_string, '%d.%m.%Y')

        for service in services:
            if service[6] == "bezahlt":
                auto_dunning = ""
                man_dunning = ""
                self.sqlstatements.update_service_dunning_auto(service, auto_dunning, man_dunning)
            else:
                accounting_date = datetime.strptime(service[1], '%d.%m.%Y')
                difference = today - accounting_date
                # Mit dem Wert difference muss geschaut werden, welche Mahnstufe in Frage kommt
                dunning = "Inkasso"
                for stage in filtered_data:
                    if int(stage[2]) >= difference.days:
                        dunning = stage[0]
                        break
                # ermittlete Stufe wird an Datenbank übergeben
                if service[6] == "abgerechnet" and service[8] == "":
                    man_dunning = "unkategorisiert"
                    self.sqlstatements.update_service_dunning_auto(service, dunning, man_dunning)
                else:
                    self.sqlstatements.update_service_dunning_auto(service, dunning, man_dunning=service[8])
        try:
            data = self.sqlstatements.get_services_for_dunning()
            self.dunning_table.row_data = data
        except:
            pass

        self.general_view_label.text = "Es wurden alle DLNe durchgerechnet"
        Clock.schedule_once(self.clear_general_view_label, 10)

    def change_dunning_status(self, instance):
        # Dropdown aus der Toolbar, wird einfach ausgewählt
        if self.dunning_table_created:
            if self.dunning_row_selected:
                result = self.sqlstatements.get_dunning_options()
                self.list = [
                    {"viewclass": "OneLineListItem",
                     "text": f"{i[0]}",
                     "on_release": lambda x=f"{i[0]}": self.update_dunning(x)}
                    for i in result]
                # DropDown wird technisch erzeugt
                self.dropdown = MDDropdownMenu(
                    items=self.list,
                    width_mult=4,
                    position="auto"
                )
                self.dropdown.caller = instance
                self.dropdown.open()
            else:
                self.general_view_label.text = "Es muss erst ein DLN ausgewählt werden"
                Clock.schedule_once(self.clear_general_view_label, 5)
        else:
            self.general_view_label.text = "Zuerst muss die Mahnungstabelle geöffnet werden"
            Clock.schedule_once(self.clear_general_view_label, 5)

    def update_dunning(self, instance):
        # update Mahnustufe und update Mahntabelle
        self.dropdown.dismiss()
        self.sqlstatements.update_service_dunning(self.dunning_info, instance)
        dunning_data = self.sqlstatements.get_services_for_dunning()
        self.dunning_table.row_data = dunning_data
        self.dunning_row_selected = False
        self.general_view_label.text = "..."

    def set_account_date(self, instance):
        status = "bezahlt"
        if self.dunning_row_selected:
            self.sqlstatements.change_status(status, self.dunning_info)
            # self.calculate wird nicht mehr automatisch gestartet,
            #self.calculate_dunning_stages(instance)
            data = self.sqlstatements.get_services_for_dunning()
            self.dunning_table.row_data = data
            self.general_view_label.text = "DLN wurde auf bezahlt gesetzt"
            Clock.schedule_once(self.clear_general_view_label, 5)
            self.dunning_row_selected = False
        else:
            self.general_view_label.text = "Es muss erst ein DLN ausgewählt werden"
            Clock.schedule_once(self.clear_general_view_label, 5)


class MyAgrar(MDApp):
    def __init__(self, mandant=None, **kwargs):
        super().__init__(**kwargs)
        self.mandant = mandant
        user_found = False
        Window.size = (1920, 1080)
        # Center the window
        Window.left = 1
        Window.top = 30
        # Window.maximize()
        os.environ['KIVY_NO_ARGS'] = '1'

        windows_username = os.environ.get('USERNAME')
        sqlstatments = SQLStatements(mandant)
        userlist = sqlstatments.get_userlist()
        for user_data in userlist:
            if user_data[4] == windows_username:
                print("Super, ich kenne dich")
                user_found = True
        if not user_found:
            Login = LoginApp(userlist=userlist)
            Login.run()

            if Login.login_successful:
                agrar = MyAgrar(mandant=mandant)
                agrar.run()

    def build(self):
        # Räume werden geladen
        self.screen_manager = ScreenManager()
        self.screen_manager.add_widget(ServicesScreen())
        self.screen_manager.add_widget(AdminScreen())

        return self.screen_manager


def get_default_mandant():
    # Über die Config wird der Standard für jeden User eingelesen
    config = configparser.ConfigParser()
    config.read('config.ini')
    return config['SETTINGS']['MANDANT']


if __name__ == "__main__":
    mandant = get_default_mandant()
    agrar = MyAgrar(mandant=mandant)
    agrar.run()
