"""
NewsApp PowerPoint Presentation Generator
Creates a professional presentation about the Android News App project
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor as RgbColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Create presentation with 16:9 aspect ratio
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Define color scheme (modern blue/purple gradient theme)
PRIMARY_COLOR = RgbColor(79, 70, 229)      # Indigo
SECONDARY_COLOR = RgbColor(139, 92, 246)   # Purple
ACCENT_COLOR = RgbColor(6, 182, 212)       # Cyan
DARK_COLOR = RgbColor(30, 41, 59)          # Slate dark
LIGHT_COLOR = RgbColor(248, 250, 252)      # Slate light
WHITE = RgbColor(255, 255, 255)
TEXT_DARK = RgbColor(15, 23, 42)

def add_gradient_background(slide):
    """Add a gradient-like background using shapes"""
    # Add background rectangle
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = DARK_COLOR
    bg.line.fill.background()
    
    # Add accent circle for visual interest
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(9), Inches(-2), Inches(6), Inches(6)
    )
    circle.fill.solid()
    circle.fill.fore_color.rgb = PRIMARY_COLOR
    circle.line.fill.background()
    
    # Add second accent
    circle2 = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(-2), Inches(4), Inches(5), Inches(5)
    )
    circle2.fill.solid()
    circle2.fill.fore_color.rgb = SECONDARY_COLOR
    circle2.line.fill.background()

def add_section_header_bg(slide):
    """Add background for section headers"""
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = LIGHT_COLOR
    bg.line.fill.background()
    
    # Add colored bar at top
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.15)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = PRIMARY_COLOR
    bar.line.fill.background()

def add_title_text(slide, text, left, top, width, height, font_size=44, bold=True, color=WHITE):
    """Add styled title text"""
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = "Segoe UI"
    return shape

def add_body_text(slide, text, left, top, width, height, font_size=18, color=TEXT_DARK):
    """Add body text"""
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.name = "Segoe UI"
    return shape

def add_bullet_points(slide, points, left, top, width, height, font_size=16, color=TEXT_DARK):
    """Add bullet points"""
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.word_wrap = True
    
    for i, point in enumerate(points):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = f"• {point}"
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = "Segoe UI"
        p.space_after = Pt(8)
    return shape

# ============================================================
# SLIDE 1: Title Slide
# ============================================================
slide1 = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
add_gradient_background(slide1)

# App icon placeholder (emoji representation)
icon_box = slide1.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.5), Inches(1.5), Inches(2), Inches(2)
)
icon_box.fill.solid()
icon_box.fill.fore_color.rgb = WHITE
icon_box.line.fill.background()

icon_text = slide1.shapes.add_textbox(Inches(5.5), Inches(1.7), Inches(2), Inches(2))
tf = icon_text.text_frame
p = tf.paragraphs[0]
p.text = "📰"
p.font.size = Pt(72)
p.alignment = PP_ALIGN.CENTER

# Title
add_title_text(slide1, "News App", Inches(0.5), Inches(4), Inches(12), Inches(1), font_size=60)

# Subtitle
add_title_text(slide1, "A Modern Android News Application", 
               Inches(0.5), Inches(5), Inches(12), Inches(0.7), font_size=28, bold=False, color=RgbColor(203, 213, 225))

# Author info
add_title_text(slide1, "Built with Kotlin & Jetpack Compose", 
               Inches(0.5), Inches(6.2), Inches(12), Inches(0.5), font_size=16, bold=False, color=RgbColor(148, 163, 184))

# ============================================================
# SLIDE 2: Overview
# ============================================================
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
add_section_header_bg(slide2)

add_title_text(slide2, "📱 App Overview", Inches(0.5), Inches(0.5), Inches(12), Inches(1), font_size=40, color=TEXT_DARK)

# Description
desc = """News App is a modern Android application that fetches and displays top headlines from around the world. 
The app connects to NewsAPI.org to download real-time news and presents them in a beautiful, 
easy-to-navigate interface built with the latest Android technologies."""

add_body_text(slide2, desc, Inches(0.5), Inches(1.5), Inches(12), Inches(1.5), font_size=18)

# Key highlights in boxes
highlights = [
    ("🌐", "Real-time News", "Live headlines from NewsAPI"),
    ("🎨", "Modern UI", "Material 3 Design System"),
    ("🌙", "Dark Mode", "Automatic theme switching"),
    ("📱", "Native Android", "Built with Kotlin")
]

for i, (emoji, title, desc) in enumerate(highlights):
    x = Inches(0.5 + (i * 3.1))
    box = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(3.5), Inches(2.9), Inches(2.5))
    box.fill.solid()
    box.fill.fore_color.rgb = WHITE
    box.line.color.rgb = RgbColor(226, 232, 240)
    
    # Emoji
    e = slide2.shapes.add_textbox(x, Inches(3.6), Inches(2.9), Inches(0.8))
    e.text_frame.paragraphs[0].text = emoji
    e.text_frame.paragraphs[0].font.size = Pt(36)
    e.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Title
    t = slide2.shapes.add_textbox(x, Inches(4.4), Inches(2.9), Inches(0.5))
    t.text_frame.paragraphs[0].text = title
    t.text_frame.paragraphs[0].font.size = Pt(16)
    t.text_frame.paragraphs[0].font.bold = True
    t.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK
    t.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Description
    d = slide2.shapes.add_textbox(x + Inches(0.1), Inches(4.9), Inches(2.7), Inches(0.8))
    d.text_frame.paragraphs[0].text = desc
    d.text_frame.paragraphs[0].font.size = Pt(12)
    d.text_frame.paragraphs[0].font.color.rgb = RgbColor(100, 116, 139)
    d.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# ============================================================
# SLIDE 3: Features
# ============================================================
slide3 = prs.slides.add_slide(prs.slide_layouts[6])
add_section_header_bg(slide3)

add_title_text(slide3, "✨ Key Features", Inches(0.5), Inches(0.5), Inches(12), Inches(1), font_size=40, color=TEXT_DARK)

features = [
    ("Live News Feed", "Connects to the internet and downloads the latest top headlines in real-time"),
    ("Beautiful Cards", "Each article displayed in a visually appealing card with image, title, description, and source"),
    ("Read Full Articles", "Tap on any news card to open the complete article in your web browser"),
    ("Modern Design", "Uses Google's latest Material 3 design system with light and dark modes"),
    ("Automatic Theme", "On Android 12+, app colors automatically match your phone's wallpaper (Dynamic Colors)")
]

for i, (title, desc) in enumerate(features):
    y = Inches(1.5 + (i * 1.1))
    
    # Feature number circle
    circle = slide3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.5), y, Inches(0.5), Inches(0.5))
    circle.fill.solid()
    circle.fill.fore_color.rgb = PRIMARY_COLOR
    circle.line.fill.background()
    
    num = slide3.shapes.add_textbox(Inches(0.5), y + Inches(0.08), Inches(0.5), Inches(0.4))
    num.text_frame.paragraphs[0].text = str(i + 1)
    num.text_frame.paragraphs[0].font.size = Pt(16)
    num.text_frame.paragraphs[0].font.bold = True
    num.text_frame.paragraphs[0].font.color.rgb = WHITE
    num.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Title
    t = slide3.shapes.add_textbox(Inches(1.2), y, Inches(11), Inches(0.4))
    t.text_frame.paragraphs[0].text = title
    t.text_frame.paragraphs[0].font.size = Pt(18)
    t.text_frame.paragraphs[0].font.bold = True
    t.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK
    
    # Description
    d = slide3.shapes.add_textbox(Inches(1.2), y + Inches(0.35), Inches(11), Inches(0.6))
    d.text_frame.paragraphs[0].text = desc
    d.text_frame.paragraphs[0].font.size = Pt(14)
    d.text_frame.paragraphs[0].font.color.rgb = RgbColor(100, 116, 139)

# ============================================================
# SLIDE 4: Technologies Used
# ============================================================
slide4 = prs.slides.add_slide(prs.slide_layouts[6])
add_section_header_bg(slide4)

add_title_text(slide4, "🛠️ Technologies Used", Inches(0.5), Inches(0.5), Inches(12), Inches(1), font_size=40, color=TEXT_DARK)

technologies = [
    ("Kotlin", "Main programming language for Android development", "🟣"),
    ("Jetpack Compose", "Google's modern toolkit for building native Android UI", "🎨"),
    ("Material 3", "Latest design system from Google for beautiful interfaces", "✨"),
    ("Retrofit", "Library that makes it easy to connect to web APIs", "🔌"),
    ("Coil", "Image loading library for efficient image display", "🖼️"),
    ("Coroutines", "Handles background tasks without freezing the app", "⚡"),
]

# Create a table-like layout
for i, (tech, desc, emoji) in enumerate(technologies):
    row = i // 2
    col = i % 2
    x = Inches(0.5 + (col * 6.3))
    y = Inches(1.5 + (row * 1.8))
    
    # Card background
    card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(6), Inches(1.5))
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = RgbColor(226, 232, 240)
    
    # Emoji
    e = slide4.shapes.add_textbox(x + Inches(0.2), y + Inches(0.3), Inches(0.6), Inches(0.6))
    e.text_frame.paragraphs[0].text = emoji
    e.text_frame.paragraphs[0].font.size = Pt(28)
    
    # Tech name
    t = slide4.shapes.add_textbox(x + Inches(0.9), y + Inches(0.25), Inches(5), Inches(0.5))
    t.text_frame.paragraphs[0].text = tech
    t.text_frame.paragraphs[0].font.size = Pt(18)
    t.text_frame.paragraphs[0].font.bold = True
    t.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK
    
    # Description
    d = slide4.shapes.add_textbox(x + Inches(0.9), y + Inches(0.7), Inches(4.9), Inches(0.7))
    d.text_frame.word_wrap = True
    d.text_frame.paragraphs[0].text = desc
    d.text_frame.paragraphs[0].font.size = Pt(12)
    d.text_frame.paragraphs[0].font.color.rgb = RgbColor(100, 116, 139)

# ============================================================
# SLIDE 5: Architecture / Project Structure
# ============================================================
slide5 = prs.slides.add_slide(prs.slide_layouts[6])
add_section_header_bg(slide5)

add_title_text(slide5, "📁 Project Structure", Inches(0.5), Inches(0.5), Inches(12), Inches(1), font_size=40, color=TEXT_DARK)

# Folder structure visualization
folders = [
    ("com.example.newsapp", "", Inches(0.5)),
    ("├── MainActivity.kt", "App entry point, sets up theme", Inches(0.7)),
    ("├── data/", "Data layer", Inches(0.7)),
    ("│   ├── NewsApi.kt", "API interface & data classes", Inches(0.9)),
    ("│   └── NewsRepository.kt", "Handles API connections", Inches(0.9)),
    ("└── ui/", "UI layer", Inches(0.7)),
    ("    ├── NewsScreen.kt", "Main screen composable", Inches(0.9)),
    ("    ├── NewsViewModel.kt", "State management", Inches(0.9)),
    ("    └── theme/", "App theming (colors, typography)", Inches(0.9)),
]

for i, (name, desc, indent) in enumerate(folders):
    y = Inches(1.4 + (i * 0.55))
    
    # File/folder name
    t = slide5.shapes.add_textbox(indent, y, Inches(5), Inches(0.5))
    t.text_frame.paragraphs[0].text = name
    t.text_frame.paragraphs[0].font.size = Pt(16)
    t.text_frame.paragraphs[0].font.name = "Consolas"
    t.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK if "/" in name or name.startswith("com") else PRIMARY_COLOR
    
    if desc:
        # Description
        d = slide5.shapes.add_textbox(Inches(6), y, Inches(6), Inches(0.5))
        d.text_frame.paragraphs[0].text = f"→ {desc}"
        d.text_frame.paragraphs[0].font.size = Pt(14)
        d.text_frame.paragraphs[0].font.color.rgb = RgbColor(100, 116, 139)

# Architecture diagram box
arch_box = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.8), Inches(12.3), Inches(1.4))
arch_box.fill.solid()
arch_box.fill.fore_color.rgb = RgbColor(239, 246, 255)
arch_box.line.color.rgb = PRIMARY_COLOR

arch_title = slide5.shapes.add_textbox(Inches(0.7), Inches(5.9), Inches(12), Inches(0.4))
arch_title.text_frame.paragraphs[0].text = "📐 Architecture Pattern: MVVM (Model-View-ViewModel)"
arch_title.text_frame.paragraphs[0].font.size = Pt(16)
arch_title.text_frame.paragraphs[0].font.bold = True
arch_title.text_frame.paragraphs[0].font.color.rgb = PRIMARY_COLOR

arch_desc = slide5.shapes.add_textbox(Inches(0.7), Inches(6.35), Inches(12), Inches(0.7))
arch_desc.text_frame.paragraphs[0].text = "UI (Compose) ← ViewModel (State) ← Repository ← API (Retrofit)"
arch_desc.text_frame.paragraphs[0].font.size = Pt(14)
arch_desc.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK

# ============================================================
# SLIDE 6: How It Works
# ============================================================
slide6 = prs.slides.add_slide(prs.slide_layouts[6])
add_section_header_bg(slide6)

add_title_text(slide6, "⚙️ How It Works", Inches(0.5), Inches(0.5), Inches(12), Inches(1), font_size=40, color=TEXT_DARK)

# Flow diagram with boxes
steps = [
    ("1", "App Launch", "MainActivity initializes\nthe NewsViewModel"),
    ("2", "Fetch News", "ViewModel calls Repository\nto get headlines"),
    ("3", "API Request", "Retrofit sends request\nto NewsAPI.org"),
    ("4", "Display", "News articles shown\nin scrollable list"),
    ("5", "Interaction", "User taps card to\nopen full article"),
]

for i, (num, title, desc) in enumerate(steps):
    x = Inches(0.3 + (i * 2.55))
    
    # Step box
    box = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.8), Inches(2.3), Inches(2.5))
    box.fill.solid()
    box.fill.fore_color.rgb = WHITE
    box.line.color.rgb = PRIMARY_COLOR
    
    # Number circle
    circle = slide6.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.85), Inches(1.5), Inches(0.6), Inches(0.6))
    circle.fill.solid()
    circle.fill.fore_color.rgb = PRIMARY_COLOR
    circle.line.fill.background()
    
    n = slide6.shapes.add_textbox(x + Inches(0.85), Inches(1.58), Inches(0.6), Inches(0.5))
    n.text_frame.paragraphs[0].text = num
    n.text_frame.paragraphs[0].font.size = Pt(18)
    n.text_frame.paragraphs[0].font.bold = True
    n.text_frame.paragraphs[0].font.color.rgb = WHITE
    n.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Title
    t = slide6.shapes.add_textbox(x + Inches(0.1), Inches(2.2), Inches(2.1), Inches(0.5))
    t.text_frame.paragraphs[0].text = title
    t.text_frame.paragraphs[0].font.size = Pt(14)
    t.text_frame.paragraphs[0].font.bold = True
    t.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK
    t.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Description
    d = slide6.shapes.add_textbox(x + Inches(0.1), Inches(2.7), Inches(2.1), Inches(1.2))
    d.text_frame.word_wrap = True
    d.text_frame.paragraphs[0].text = desc
    d.text_frame.paragraphs[0].font.size = Pt(11)
    d.text_frame.paragraphs[0].font.color.rgb = RgbColor(100, 116, 139)
    d.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Arrow (except last)
    if i < 4:
        arrow = slide6.shapes.add_textbox(x + Inches(2.3), Inches(2.8), Inches(0.3), Inches(0.5))
        arrow.text_frame.paragraphs[0].text = "→"
        arrow.text_frame.paragraphs[0].font.size = Pt(24)
        arrow.text_frame.paragraphs[0].font.color.rgb = PRIMARY_COLOR

# Data flow section
flow_box = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(4.8), Inches(12.3), Inches(2.4))
flow_box.fill.solid()
flow_box.fill.fore_color.rgb = RgbColor(254, 252, 232)
flow_box.line.color.rgb = RgbColor(250, 204, 21)

flow_title = slide6.shapes.add_textbox(Inches(0.7), Inches(4.95), Inches(12), Inches(0.4))
flow_title.text_frame.paragraphs[0].text = "📊 Data Flow"
flow_title.text_frame.paragraphs[0].font.size = Pt(18)
flow_title.text_frame.paragraphs[0].font.bold = True
flow_title.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK

flow_items = [
    "• NewsViewModel uses StateFlow to hold the list of articles and loading state",
    "• NewsScreen observes the state using collectAsState() and reacts to changes",
    "• Coroutines handle the API call on a background thread, preventing UI freezes",
    "• Loading indicator shown while fetching, articles displayed once data arrives"
]

for i, item in enumerate(flow_items):
    f = slide6.shapes.add_textbox(Inches(0.7), Inches(5.45 + (i * 0.4)), Inches(12), Inches(0.4))
    f.text_frame.paragraphs[0].text = item
    f.text_frame.paragraphs[0].font.size = Pt(13)
    f.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK

# ============================================================
# SLIDE 7: Code Highlights
# ============================================================
slide7 = prs.slides.add_slide(prs.slide_layouts[6])
add_section_header_bg(slide7)

add_title_text(slide7, "💻 Code Highlights", Inches(0.5), Inches(0.5), Inches(12), Inches(1), font_size=40, color=TEXT_DARK)

# Code box 1 - NewsApi
code_box1 = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), Inches(6), Inches(2.7))
code_box1.fill.solid()
code_box1.fill.fore_color.rgb = RgbColor(30, 41, 59)
code_box1.line.fill.background()

code1_title = slide7.shapes.add_textbox(Inches(0.7), Inches(1.4), Inches(5.5), Inches(0.4))
code1_title.text_frame.paragraphs[0].text = "NewsApi.kt - API Interface"
code1_title.text_frame.paragraphs[0].font.size = Pt(12)
code1_title.text_frame.paragraphs[0].font.color.rgb = ACCENT_COLOR

code1 = """interface NewsApi {
    @GET("top-headlines")
    suspend fun getTopHeadlines(
        @Query("country") country: String,
        @Query("apiKey") apiKey: String
    ): NewsResponse
}"""

code1_text = slide7.shapes.add_textbox(Inches(0.7), Inches(1.75), Inches(5.5), Inches(2.1))
code1_text.text_frame.word_wrap = True
code1_text.text_frame.paragraphs[0].text = code1
code1_text.text_frame.paragraphs[0].font.size = Pt(11)
code1_text.text_frame.paragraphs[0].font.name = "Consolas"
code1_text.text_frame.paragraphs[0].font.color.rgb = RgbColor(226, 232, 240)

# Code box 2 - ViewModel
code_box2 = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.3), Inches(6), Inches(2.7))
code_box2.fill.solid()
code_box2.fill.fore_color.rgb = RgbColor(30, 41, 59)
code_box2.line.fill.background()

code2_title = slide7.shapes.add_textbox(Inches(7), Inches(1.4), Inches(5.5), Inches(0.4))
code2_title.text_frame.paragraphs[0].text = "NewsViewModel.kt - State Management"
code2_title.text_frame.paragraphs[0].font.size = Pt(12)
code2_title.text_frame.paragraphs[0].font.color.rgb = ACCENT_COLOR

code2 = """class NewsViewModel : ViewModel() {
    private val _articles = 
        MutableStateFlow<List<Article>>(emptyList())
    val articles: StateFlow<List<Article>> = 
        _articles.asStateFlow()
    
    fun fetchNews() { ... }
}"""

code2_text = slide7.shapes.add_textbox(Inches(7), Inches(1.75), Inches(5.5), Inches(2.1))
code2_text.text_frame.word_wrap = True
code2_text.text_frame.paragraphs[0].text = code2
code2_text.text_frame.paragraphs[0].font.size = Pt(11)
code2_text.text_frame.paragraphs[0].font.name = "Consolas"
code2_text.text_frame.paragraphs[0].font.color.rgb = RgbColor(226, 232, 240)

# Code box 3 - Compose UI
code_box3 = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(4.2), Inches(12.3), Inches(3))
code_box3.fill.solid()
code_box3.fill.fore_color.rgb = RgbColor(30, 41, 59)
code_box3.line.fill.background()

code3_title = slide7.shapes.add_textbox(Inches(0.7), Inches(4.3), Inches(11.8), Inches(0.4))
code3_title.text_frame.paragraphs[0].text = "NewsScreen.kt - Jetpack Compose UI"
code3_title.text_frame.paragraphs[0].font.size = Pt(12)
code3_title.text_frame.paragraphs[0].font.color.rgb = ACCENT_COLOR

code3 = """@Composable
fun NewsScreen(viewModel: NewsViewModel) {
    val articles by viewModel.articles.collectAsState()
    
    LazyColumn {
        items(articles) { article ->
            ArticleItem(article)    // Card with image, title, description
        }
    }
}"""

code3_text = slide7.shapes.add_textbox(Inches(0.7), Inches(4.65), Inches(11.8), Inches(2.4))
code3_text.text_frame.word_wrap = True
code3_text.text_frame.paragraphs[0].text = code3
code3_text.text_frame.paragraphs[0].font.size = Pt(13)
code3_text.text_frame.paragraphs[0].font.name = "Consolas"
code3_text.text_frame.paragraphs[0].font.color.rgb = RgbColor(226, 232, 240)

# ============================================================
# SLIDE 8: App Theming
# ============================================================
slide8 = prs.slides.add_slide(prs.slide_layouts[6])
add_section_header_bg(slide8)

add_title_text(slide8, "🎨 App Theming", Inches(0.5), Inches(0.5), Inches(12), Inches(1), font_size=40, color=TEXT_DARK)

# Light mode box
light_box = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.5), Inches(4), Inches(4))
light_box.fill.solid()
light_box.fill.fore_color.rgb = WHITE
light_box.line.color.rgb = RgbColor(226, 232, 240)

light_header = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.5), Inches(4), Inches(0.6))
light_header.fill.solid()
light_header.fill.fore_color.rgb = PRIMARY_COLOR
light_header.line.fill.background()

light_title = slide8.shapes.add_textbox(Inches(0.7), Inches(1.55), Inches(3.6), Inches(0.5))
light_title.text_frame.paragraphs[0].text = "☀️ Light Mode"
light_title.text_frame.paragraphs[0].font.size = Pt(18)
light_title.text_frame.paragraphs[0].font.bold = True
light_title.text_frame.paragraphs[0].font.color.rgb = WHITE

light_desc = slide8.shapes.add_textbox(Inches(0.7), Inches(2.3), Inches(3.6), Inches(2.5))
light_desc.text_frame.word_wrap = True
tf = light_desc.text_frame
p = tf.paragraphs[0]
p.text = "• Bright, clean appearance\n• White backgrounds\n• High contrast text\n• Perfect for daytime use"
p.font.size = Pt(14)
p.font.color.rgb = TEXT_DARK

# Dark mode box
dark_box = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.75), Inches(1.5), Inches(4), Inches(4))
dark_box.fill.solid()
dark_box.fill.fore_color.rgb = RgbColor(30, 41, 59)
dark_box.line.color.rgb = RgbColor(51, 65, 85)

dark_header = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.75), Inches(1.5), Inches(4), Inches(0.6))
dark_header.fill.solid()
dark_header.fill.fore_color.rgb = SECONDARY_COLOR
dark_header.line.fill.background()

dark_title = slide8.shapes.add_textbox(Inches(4.95), Inches(1.55), Inches(3.6), Inches(0.5))
dark_title.text_frame.paragraphs[0].text = "🌙 Dark Mode"
dark_title.text_frame.paragraphs[0].font.size = Pt(18)
dark_title.text_frame.paragraphs[0].font.bold = True
dark_title.text_frame.paragraphs[0].font.color.rgb = WHITE

dark_desc = slide8.shapes.add_textbox(Inches(4.95), Inches(2.3), Inches(3.6), Inches(2.5))
dark_desc.text_frame.word_wrap = True
tf = dark_desc.text_frame
p = tf.paragraphs[0]
p.text = "• Eye-friendly dark interface\n• Reduced eye strain\n• Battery saving on OLED\n• Perfect for nighttime"
p.font.size = Pt(14)
p.font.color.rgb = RgbColor(226, 232, 240)

# Dynamic colors box
dyn_box = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9), Inches(1.5), Inches(4), Inches(4))
# Gradient-like effect with multiple colors
dyn_box.fill.solid()
dyn_box.fill.fore_color.rgb = RgbColor(254, 215, 170)
dyn_box.line.color.rgb = RgbColor(251, 146, 60)

dyn_header = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9), Inches(1.5), Inches(4), Inches(0.6))
dyn_header.fill.solid()
dyn_header.fill.fore_color.rgb = ACCENT_COLOR
dyn_header.line.fill.background()

dyn_title = slide8.shapes.add_textbox(Inches(9.2), Inches(1.55), Inches(3.6), Inches(0.5))
dyn_title.text_frame.paragraphs[0].text = "🌈 Dynamic Colors"
dyn_title.text_frame.paragraphs[0].font.size = Pt(18)
dyn_title.text_frame.paragraphs[0].font.bold = True
dyn_title.text_frame.paragraphs[0].font.color.rgb = WHITE

dyn_desc = slide8.shapes.add_textbox(Inches(9.2), Inches(2.3), Inches(3.6), Inches(2.5))
dyn_desc.text_frame.word_wrap = True
tf = dyn_desc.text_frame
p = tf.paragraphs[0]
p.text = "• Android 12+ feature\n• Colors from wallpaper\n• Personalized experience\n• Material You design"
p.font.size = Pt(14)
p.font.color.rgb = TEXT_DARK

# Implementation note
impl_box = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.8), Inches(12.3), Inches(1.4))
impl_box.fill.solid()
impl_box.fill.fore_color.rgb = RgbColor(240, 253, 244)
impl_box.line.color.rgb = RgbColor(34, 197, 94)

impl_text = slide8.shapes.add_textbox(Inches(0.7), Inches(6), Inches(12), Inches(1))
impl_text.text_frame.word_wrap = True
impl_text.text_frame.paragraphs[0].text = "✅ Implementation: Theme automatically switches based on system settings. Uses dynamicDarkColorScheme() and dynamicLightColorScheme() on Android 12+, with fallback to defined color schemes on older versions."
impl_text.text_frame.paragraphs[0].font.size = Pt(14)
impl_text.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK

# ============================================================
# SLIDE 9: Future Improvements
# ============================================================
slide9 = prs.slides.add_slide(prs.slide_layouts[6])
add_section_header_bg(slide9)

add_title_text(slide9, "🚀 Future Improvements", Inches(0.5), Inches(0.5), Inches(12), Inches(1), font_size=40, color=TEXT_DARK)

improvements = [
    ("🔍", "Search", "Let users search for news on specific topics"),
    ("📂", "Categories", "Filter news by sports, tech, business, etc."),
    ("⭐", "Favorites", "Save articles to read later"),
    ("📴", "Offline Mode", "Store articles for reading without internet"),
    ("🔄", "Pull to Refresh", "Swipe down to get latest headlines"),
    ("📤", "Share", "Share interesting articles with friends"),
    ("🌍", "Multi-Country", "Choose news from different countries"),
    ("🔔", "Notifications", "Get alerts for breaking news"),
]

for i, (emoji, title, desc) in enumerate(improvements):
    row = i // 4
    col = i % 4
    x = Inches(0.4 + (col * 3.2))
    y = Inches(1.4 + (row * 2.8))
    
    # Card
    card = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(3), Inches(2.5))
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = RgbColor(226, 232, 240)
    
    # Emoji
    e = slide9.shapes.add_textbox(x, y + Inches(0.2), Inches(3), Inches(0.7))
    e.text_frame.paragraphs[0].text = emoji
    e.text_frame.paragraphs[0].font.size = Pt(32)
    e.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Title
    t = slide9.shapes.add_textbox(x, y + Inches(0.9), Inches(3), Inches(0.5))
    t.text_frame.paragraphs[0].text = title
    t.text_frame.paragraphs[0].font.size = Pt(16)
    t.text_frame.paragraphs[0].font.bold = True
    t.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK
    t.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Description
    d = slide9.shapes.add_textbox(x + Inches(0.15), y + Inches(1.4), Inches(2.7), Inches(0.9))
    d.text_frame.word_wrap = True
    d.text_frame.paragraphs[0].text = desc
    d.text_frame.paragraphs[0].font.size = Pt(11)
    d.text_frame.paragraphs[0].font.color.rgb = RgbColor(100, 116, 139)
    d.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# ============================================================
# SLIDE 10: Thank You
# ============================================================
slide10 = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_background(slide10)

# Thank you text
add_title_text(slide10, "Thank You!", Inches(0.5), Inches(2.5), Inches(12.3), Inches(1.2), font_size=72)

# Subtitle
add_title_text(slide10, "Questions?", Inches(0.5), Inches(4), Inches(12.3), Inches(0.8), 
               font_size=36, bold=False, color=RgbColor(203, 213, 225))

# Summary box
summary_box = slide10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.5), Inches(5.2), Inches(6.3), Inches(1.8))
summary_box.fill.solid()
summary_box.fill.fore_color.rgb = RgbColor(255, 255, 255)
summary_box.fill.fore_color.rgb._alpha = 0.1  # Semi-transparent
summary_box.line.fill.background()

summary_text = slide10.shapes.add_textbox(Inches(3.7), Inches(5.4), Inches(5.9), Inches(1.5))
summary_text.text_frame.word_wrap = True
tf = summary_text.text_frame
p = tf.paragraphs[0]
p.text = "📰 News App\nBuilt with Kotlin, Jetpack Compose & Material 3\nReal-time headlines from NewsAPI.org"
p.font.size = Pt(16)
p.font.color.rgb = TEXT_DARK
p.alignment = PP_ALIGN.CENTER

# Save the presentation
output_path = r"c:\Users\LenovoPC\Desktop\first_app\NewsApp_Presentation.pptx"
prs.save(output_path)
print(f"\n✅ Presentation saved successfully!")
print(f"📁 Location: {output_path}")
print(f"\n📊 Presentation contains {len(prs.slides)} slides:")
print("   1. Title Slide")
print("   2. App Overview")
print("   3. Key Features")
print("   4. Technologies Used")
print("   5. Project Structure")
print("   6. How It Works")
print("   7. Code Highlights")
print("   8. App Theming")
print("   9. Future Improvements")
print("   10. Thank You / Q&A")
