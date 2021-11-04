from botcity.core import DesktopBot
import xlsxwriter
import os

class Bot(DesktopBot):
    def action(self, execution=None):
        self.load()

        movies = self.getMovies()
        movies_review = self.searchMovieRating(movies)
        saveMoviesReview(movies_review)

    # Getting the 3 most popular movies from rpa challenge website
    def getMovies(self):
        self.browse("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe")
        self.wait(4000)
        self.paste("https://www.rpachallenge.com", 1000)
        self.enter()

        movie_list = []

        if not self.find( "movie_search", matching=0.97, waiting_time=20000):
            self.not_found("movie_search")
        self.click()
        
        if not self.find( "get_movies", matching=0.97, waiting_time=10000):
            self.not_found("get_movies")
        self.click()
        
        self.wait(1000)
        self.control_a(1000)
        self.control_c(1000)
        page_content = self.get_clipboard()
        self.click_relative(0, -50)
        
        # Parsing the page content and extracting only the movies name
        data = page_content.split('\n')
        for line in data:
            if line.__contains__('commentdelete'):
                movie_list.append(line.replace('commentdelete', ''))
        
        print('\nMovies => ' + str(movie_list) + '\n')

        return movie_list

    # Searching the movies rating in rotten tomatoes website
    def searchMovieRating(self, movie_list):
        self.control_t()
        self.wait(4000)
        self.paste("https://www.rottentomatoes.com")
        self.enter()
        self.wait(5000)

        if not self.find( "search_movies", matching=0.97, waiting_time=10000):
                self.not_found("search_movies")
        self.click()
        
        movie_review_list = []

        # For every movie, collect the page content and the reviews score
        for movie in movie_list:
            self.wait(1000)
            self.paste(str(movie), 3000)
            self.tab(1000)
            self.tab(1000)
            self.tab(1000)
            self.enter()

            self.wait(10000)
            self.control_a(1000)
            self.control_c(1000)
            movie_page = self.get_clipboard()
            self.wait(1000)
            self.click_at(200,200)

            # Extracting the reviewers score and the audience score
            start_index = str(movie_page).index("View All")
            end_index = str(movie_page).index("AUDIENCE SCORE")
            rating_data = str(movie_page[start_index : end_index]).split('\n')       
            reviewers_score = rating_data[4].replace('\r', '')
            audience_score = rating_data[7].replace('\r', '')

            if not reviewers_score.__contains__("%"):
                reviewers_score = " - "
            if not audience_score.__contains__("%"):
                audience_score = " - "

            movie_review = []
            movie_review.append(str(movie).replace('\r', ''))
            movie_review.append(reviewers_score)
            movie_review.append(audience_score)

            print(str(movie_review))
            movie_review_list.append(movie_review)

            if not self.find( "search_movies2", matching=0.97, waiting_time=10000):
                self.not_found("search_movies2")
            self.click()

        self.alt_f4()
        return movie_review_list

    def load(self):
        self.add_image("movie_search", self.get_resource_abspath("movie_search.png"))
        self.add_image("get_movies", self.get_resource_abspath("get_movies.png"))
        self.add_image("search_movies", self.get_resource_abspath("search_movies.png"))
        self.add_image("search_movies2", self.get_resource_abspath("search_movies2.png"))

    def not_found(self, label):
        print(f"Element not found: {label}")

# Updating the sheet with the movies name and review
def saveMoviesReview(data):
    workbook = xlsxwriter.Workbook('popular_movies_review.xlsx')
    bold = workbook.add_format({'bold': True})
    worksheet = workbook.add_worksheet()
    worksheet.set_column(0, 2, 40)

    worksheet.write('A1', 'Movie Name', bold)
    worksheet.write('B1', 'Review Score', bold)
    worksheet.write('C1', 'Audience Score', bold)

    row = 1
    col = 0
    for review in data:
        worksheet.write_row(row, col, review)
        row += 1

    workbook.close()
    os.system(" start EXCEL.EXE popular_movies_review.xlsx")

if __name__ == '__main__':
    Bot.main()
