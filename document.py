import calendar
import click
import locale
import yaml
from datetime import date
from mailmerge import MailMerge
from os import path


locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
CONFIG = "config.yml"
TEMPLATE_1 = "cuenta-de-cobro-v1.docx"
TEMPLATE_2 = "cuenta-de-cobro-v2.docx"


def load_config():
    with open(CONFIG, "r") as ymlfile:
        cfg = yaml.load(ymlfile, Loader=yaml.FullLoader)
    return cfg


def get_date(year=None, month=None):
    today = date.today()
    if year and month:
        today = date(year, month, calendar.monthrange(year, month)[1])
    days = calendar.monthrange(
        int(today.strftime("%Y")), 
        int(today.strftime("%m"))
    )[1]
    formated_date = today.strftime(r"%d de %B de %Y")
    formated_date_start = today.replace(day=1).strftime("%d de %B de %Y")
    formated_date_end = today.replace(day=days).strftime("%d de %B de %Y")
    return (formated_date, formated_date_start, formated_date_end)


def create_document(cfg, output="output.docx", cfg_date=False):
    template_docx = TEMPLATE_1
    if cfg_date:
        template_docx = TEMPLATE_2

    document = MailMerge(template_docx)
    # print(document.get_merge_fields())
    section_user = cfg["user"]
    section_bank = cfg["bank"]
    section_target = cfg["target"]
    section_service = cfg["service"]
    section_custom_date = cfg["custom_date"]

    if cfg_date:
        document.merge(
            amount=section_service["amount"],
            bankName=section_bank["bankName"],
            bankNum=str(section_bank["bankNum"]),
            business=section_target["business"],
            name=section_user["name"],
            nameUpper=section_user["name"].upper(),
            nit=section_target["nit"],
            nid=str(section_user["nid"]),
            nidCity=section_user["nidCity"],
            day=str(section_custom_date["day"]),
            dayStart=str(section_custom_date["dayStart"]),
            dayEnd=str(section_custom_date["dayEnd"]),
            month=section_custom_date["month"],
            monthStart=section_custom_date["monthStart"],
            monthEnd=section_custom_date["monthEnd"],
            year=str(section_custom_date["year"]),
            yearStart=str(section_custom_date["yearStart"]),
            yearEnd=str(section_custom_date["yearEnd"]),
        )
    else:
        formated_date, formated_date_start, formated_date_end = get_date()
        document.merge(
            amount=section_service["amount"],
            bankName=section_bank["bankName"],
            bankNum=str(section_bank["bankNum"]),
            business=section_target["business"],
            name=section_user["name"],
            nameUpper=section_user["name"].upper(),
            nit=section_target["nit"],
            nid=str(section_user["nid"]),
            nidCity=section_user["nidCity"],
            currentDate=formated_date,
            dateStart=formated_date_start,
            dateEnd=formated_date_end,
        )
        
    document.write(output)
    document.close()


def batch_of_documents(cfg, year, output_dir=""):
    template_docx = TEMPLATE_1
    section_user = cfg["user"]
    section_bank = cfg["bank"]
    section_target = cfg["target"]
    section_service = cfg["service"]

    def dates():
        for m in range(1,13):
            yield get_date(year=year, month=m)

    for r in dates():
        formated_date, formated_date_start, formated_date_end = r
        output_file = path.join(output_dir, f"{formated_date}.docx")
        with MailMerge(template_docx) as document:
            document.merge(
                amount=section_service["amount"],
                bankName=section_bank["bankName"],
                bankNum=str(section_bank["bankNum"]),
                business=section_target["business"],
                name=section_user["name"],
                nameUpper=section_user["name"].upper(),
                nit=section_target["nit"],
                nid=str(section_user["nid"]),
                nidCity=section_user["nidCity"],
                currentDate=formated_date,
                dateStart=formated_date_start,
                dateEnd=formated_date_end,
            )
            document.write(output_file)


@click.group()
def cli():
    pass


@cli.command()
@click.option("--output", default="output.docx", help="Output filename")
@click.option("--custom-date", is_flag=True, help="Use date from config.yml")
def single(output, custom_date):
    conf = load_config()
    create_document(cfg=conf, output=output, cfg_date=custom_date)


@cli.command()
@click.option("--year", required=True, type=int, help="")
@click.option("--output-dir", default="", help="")
def batch(year, output_dir):
    conf = load_config()
    batch_of_documents(cfg=conf, year=year, output_dir=output_dir)


if __name__ == "__main__":
    cli()
