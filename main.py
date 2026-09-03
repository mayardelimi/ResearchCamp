from src.tools.tools import web_search , scrape_url


result = scrape_url.invoke({
    "url": "https://www.nature.com/articles/s41586-020-2649-2"
})

print(result)