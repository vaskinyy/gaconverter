import click

from gaconverter.converter import Converter

@click.group()
def cli():
    pass


@click.command()
def alleles():
    click.echo('Dropped the database')

@click.command()
@click.argument('rtf_path', nargs=1)
@click.argument('xlsx_path', nargs=1)
def convert(rtf_path, xlsx_path):
    converter = Converter(rtf_path, xlsx_path)
    converter.run()

cli.add_command(convert)
cli.add_command(alleles)

if __name__ == '__main__':
    cli()
