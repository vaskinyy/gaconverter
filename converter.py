import click

from gaconverter.converter import Converter


@click.command()
@click.argument('rtf_path', nargs=1)
@click.argument('xlsx_path', nargs=1)
def convert(rtf_path, xlsx_path):
    converter = Converter(rtf_path, xlsx_path)
    converter.run()

if __name__ == '__main__':
    convert()
