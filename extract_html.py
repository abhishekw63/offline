import os
from bs4 import BeautifulSoup

# Load HTML
with open('downloaded_homepage.html', 'r', encoding='utf-8') as f:
    html_content = f.read()

soup = BeautifulSoup(html_content, 'html.parser')

# We'll extract basic parts and clean up
# Keep <head> (CSS/JS links might be somewhat broken, but we'll try to keep styling)
# Keep <header>, <main> without products, <footer>
# Remove <cart-drawer>, checkout forms, etc.

# Let's target some of the obvious e-commerce components to remove
for selector in [
    'cart-drawer', 'cart-drawer-overlay', '#CartDrawer', '.cart-drawer',
    'form[action="/cart/add"]', 'form[action="/cart"]', '.product-card-wrapper',
    '.collection', '.quick-add', '.quantity-popover', '.cart-notification',
    '.shopify-section-header-cart', '.header__icon--cart', 'form[action*="/checkout"]',
    '.predictive-search', '.drawer', '#MainContent', '.shopify-section-collection-list',
    '.shopify-section-featured-collection', '#shopify-section-template--21508204232938__featured_collection_r3d7yZ',
]:
    for element in soup.select(selector):
        element.decompose()

# The entire <main> block might be full of product sliders, let's keep it mostly empty
# but preserve the basic structure
main_content = soup.find('main', id='MainContent')
if main_content:
    main_content.clear()
else:
    main_content = soup.find('main')
    if main_content:
        main_content.clear()

# If <main> is removed by `#MainContent` above, we should create a new main element
if not soup.find('main'):
    main_node = soup.new_tag('main', id="MainContent", attrs={'class': 'content-for-layout focus-none'})
    header = soup.find('header')
    if header:
        header.insert_after(main_node)
    else:
        body = soup.find('body')
        if body:
            body.insert(0, main_node)

# Create the login block that will go into <main>
main_node = soup.find('main')
if main_node:
    # We will inject some basic Django template structure here
    django_logic = """
    <div class="page-width page-width--narrow section-{{ section.id }}-padding" style="margin-top: 50px; margin-bottom: 50px; text-align: center;">
      {% if not user.is_authenticated %}
        <h1 class="main-page-title page-title h0">Login</h1>
        <div class="customer login">
            <form method="post" action="{% url 'login' %}">
                {% csrf_token %}
                {{ form.as_p }}
                <button type="submit" class="button">Log In</button>
            </form>
        </div>
      {% else %}
        <h1 class="main-page-title page-title h0">Welcome, {{ user.username }}</h1>
        <p>You are successfully logged in to the RENEE Warehouse application.</p>
        <a href="{% url 'index' %}" class="button">Go to GT Mass Dump Generator</a>
        <br><br>
        <form method="post" action="{% url 'logout' %}">
            {% csrf_token %}
            <button type="submit" class="button button--secondary">Log Out</button>
        </form>
      {% endif %}
    </div>
    """
    main_node.append(BeautifulSoup(django_logic, 'html.parser'))

# Remove shopify script tags that might cause errors
for script in soup.find_all('script'):
    if script.get('src') and 'shopify' in script.get('src'):
        script.decompose()

# Add standard {% load static %} if needed, but for now we rely on the CDN links
# already in the downloaded HTML.
output_html = str(soup)

# Prepend {% load static %} just in case
output_html = "{% load static %}\n" + output_html

with open('offline/templates/offline/home.html', 'w', encoding='utf-8') as f:
    f.write(output_html)

print("HTML extracted and cleaned. Saved to offline/templates/offline/home.html")
