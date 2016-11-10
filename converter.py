import click

from gaconverter.converter import Converter
from gaconverter.allele_extractor import Extractor

@click.group()
def cli():
    pass


@click.command()
@click.argument('xlsx_path', nargs=1)
def allele(xlsx_path):
    extractor = Extractor(xlsx_path)
    extractor.run()


@click.command()
@click.argument('rtf_path', nargs=1)
@click.argument('xlsx_path', nargs=1)
def convert(rtf_path, xlsx_path):
    converter = Converter(rtf_path, xlsx_path)
    converter.run()

cli.add_command(convert)
cli.add_command(allele)

if __name__ == '__main__':
    cli()
