from Google import Create_Service
from googleapiclient.http import MediaFileUpload
import json
import os
from datetime import datetime


class Module_drive():

    def __init__(self):
        api_version = "v3"
        api_name = "drive"
        scopes = ["https://www.googleapis.com/auth/drive"]
        self.service = Create_Service("client_secrets.json", api_name, api_version, scopes)

    def create_folder(self, new_folder_name, *parent):
        """
        Cria uma nova pasta no Google Drive.

        :param str new_folder_name: Nome da pasta a ser criada.
        :param str parent: Nome da pasta mãe (acima da pasta que será criada).
            """
        file_metadata = {
            'name': new_folder_name,
            'mimeType': "application/vnd.google-apps.folder",
            'parents': list(parent)
        }
        self.service.files().create(body=file_metadata).execute()

    @staticmethod
    def find_root(client):
        """
        Encontra a pasta do Google Drive no arquivo de configurações
        God has forsaken us!

        :param str client: Um dos clientes em config.json (Bermudes_DF, Bermudes_RJ, JG, TAVAD)
        :return str
        """
        client_folder = ""
        f = open("./config/config.json")
        infos = json.load(f)
        for k, v in infos.items():
            for k1, v1 in v.items():
                if k1 == client:
                    for k2, v2 in v1.items():
                        if k2 == "drive_path":
                            client_folder = v2
        return client_folder

    @staticmethod
    def find_mime(file):
        """
        Encontrar o metadata correto do documento para envio ao Google Drive.
        Usa o MIME.json para identificar o MIME correto.

        :param str file: Documento o qual se precisa descobrir o MIME
        :return str
        """
        file_mime = ""
        f = open("./config/MIME.json")
        file_extension = f".{file.split('.')[-1]}"
        mimes = json.load(f)
        for key, value in mimes.items():
            if key == file_extension:
                file_mime = value
                break
            else:
                file_mime = ""
        f.close()
        return file_mime

    @staticmethod
    def find_date():
        """Encontra as data para localizar a pasta no drive"""
        try:  # Windows
            month = datetime.today().strftime("%#m")
        except:  # Linux
            month = datetime.today().strftime("%-m")

        year = datetime.today().strftime('%Y')

        return month, year

    def upload_file(self, file, *parent):
        """
        Realiza o upload de um documento para o Google Drive

        :param str file: Path do documento que será enviado ao Google Drive
        :param str parent: ID da pasta onde o documento será salvo no Google Drive
        """
        mime = self.find_mime(file)
        head, tail = os.path.split(file)

        file_metadata = {
            'name': tail,
            'parents': list(parent)
        }

        media = MediaFileUpload(file, mimetype=mime)
        self.service.files().create(body=file_metadata, media_body=media, fields="id").execute()

    def find_folder(self, root):
        """
        Localiza a pasta onde o documento será salvo. A pasta final é: raiz -> ano -> mês

        :param str root: Pasta onde a busca será feita. Essa é a pasta mãe da pasta com ano e mês. Está em config.json
        :return str
        """
        child = ""
        final_folder = ""
        month, year = self.find_date()

        year_querry = f"parents = '{root}'"
        year_children = self.service.files().list(includeItemsFromAllDrives=True, supportsAllDrives=True,
                                                  q=year_querry).execute()
        year_folders = year_children.get('files')

        # Verifica se há uma pasta para o ano presente. Se não houver, cria a pasta.
        needs_year = True
        for i in range(len(year_folders)):
            if year in year_folders[i].values():
                needs_year = False

        if needs_year:
            self.create_folder(year, root)
            year_children = self.service.files().list(includeItemsFromAllDrives=True, supportsAllDrives=True,
                                                      q=year_querry).execute()
            year_folders = year_children.get('files')

        for year_folder in year_folders:
            if year_folder.get("name") == year:
                child = year_folder.get("id")

        month_querry = f"parents = '{child}'"
        month_children = self.service.files().list(includeItemsFromAllDrives=True, supportsAllDrives=True,
                                                   q=month_querry).execute()
        month_folders = month_children.get('files')

        # Verifica se há uma pasta para o mês presente. Se não houver, cria a pasta.
        needs_month = True
        for i in range(len(month_folders)):
            if month in month_folders[i].values():
                needs_month = False

        if needs_month:
            long_date = month
            f = open("./config/dates.json")
            data = json.load(f)
            for key, value in data.items():
                if key == month:
                    long_date = value
                    break
            f.close()
            self.create_folder(long_date, child)
            month_children = self.service.files().list(includeItemsFromAllDrives=True, supportsAllDrives=True,
                                                       q=month_querry).execute()
            month_folders = month_children.get('files')

        for month_folder in month_folders:
            if month in month_folder.get("name"):
                final_folder = month_folder.get("id")

        return final_folder

    def run_processes(self, file, client):
        """
        Roda o processo de upload de documento

        :param str file: Caminho do documento que será enviado para o drive.
        :param str client: Um dos clientes em config.json (Bermudes_DF, Bermudes_RJ, JG, TAVAD)
        """
        root_id = self.find_root(client)
        folder = self.find_folder(root_id)
        self.upload_file(file, folder)


if __name__ == "__main__":
    drive = Module_drive()
    for file in os.listdir("./reports"):
        drive.run_processes(f"./reports/{file}", "tests")

    # TODO Excluir o documento na máquina após ser enviado para o Drive. Apagar apenas quando o documento for enviado
    #  com sucesso.
