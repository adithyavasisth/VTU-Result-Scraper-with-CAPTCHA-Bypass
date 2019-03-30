import zipfile

from flask import Flask, send_from_directory
from flask_restful import Resource, Api

import scraper

app = Flask(__name__)
api = Api(app)


class ResultScraper(Resource):
    def get(self, college, year, branch, low, high, semc):
        scraper.scrape(college, year, branch, low, high, semc)
        filename = 'ExcelFiles/' + '1' + college + year + branch + low + '-' + high
        extension = '.xls'
        zipf = zipfile.ZipFile('Results-Excel.zip', 'w', zipfile.ZIP_DEFLATED)
        files = [filename + extension, filename + 'GPA' + extension, filename + 'RANK' + extension]
        for file in files:
            zipf.write(file)
        zipf.close()
        return send_from_directory('', 'Results-Excel.zip')


api.add_resource(ResultScraper, '/scrape/<college>/<year>/<branch>/<low>/<high>/<semc>')

if __name__ == '__main__':
    app.run()
