# 📰 News App

A modern Android news application that fetches and displays top headlines from around the world. Built with Kotlin and Jetpack Compose, this app provides a clean and intuitive way to stay updated with the latest news.

---

## ✨ Features

- **Live News Feed** - The app connects to the internet and downloads the latest top headlines in real-time
- **Beautiful Cards** - Each news article is displayed in a visually appealing card format with the article image, title, description, and source
- **Read Full Articles** - Simply tap on any news card to open the complete article in your web browser
- **Modern Design** - Uses Google's latest Material 3 design system with support for both light and dark modes
- **Automatic Theme** - On Android 12 and above, the app colors automatically match your phone's wallpaper

---

## 🛠️ Technologies Used

| Technology | Purpose |
|------------|---------|
| **Kotlin** | The main programming language for Android development |
| **Jetpack Compose** | Google's modern toolkit for building native Android UI |
| **Material 3** | The latest design system from Google for beautiful interfaces |
| **Retrofit** | A library that makes it easy to connect to web APIs and download data |
| **Coil** | An image loading library that efficiently downloads and displays article images |
| **Coroutines** | Handles background tasks like downloading news without freezing the app |

---

## 📱 How the App Works

### Fetching the News

When you open the app, it automatically connects to **NewsAPI.org** - a free service that provides news headlines from various sources around the world. The app requests the top headlines for the United States and receives the data in a structured format.

### Displaying Articles

Once the news data is downloaded, the app organizes it into a scrollable list. Each article appears as a card showing:

- **Image** - The featured photo from the article (if available)
- **Title** - The headline of the news article in bold
- **Description** - A brief summary of what the article is about (up to 3 lines)
- **Source** - The name of the news outlet that published the article

### Opening Full Articles

When you tap on any news card, the app opens your default web browser and takes you directly to the original article on the publisher's website, where you can read the complete story.

### Loading State

While the app is fetching news from the internet, it shows a spinning loading indicator in the center of the screen. Once the data is ready, the indicator disappears and the news articles appear.

---

## � Project Organization

The project is organized into logical folders to keep the code clean and maintainable:

### Main Files

- **MainActivity.kt** - The starting point of the app that sets up the screen and theme

### Data Folder

This folder contains everything related to getting data from the internet:

- **NewsApi.kt** - Defines what data we want from the news service and how articles are structured
- **NewsRepository.kt** - Handles the actual connection to the internet and downloading of news

### UI Folder

This folder contains everything the user sees and interacts with:

- **NewsScreen.kt** - The main screen layout with the list of news articles
- **NewsViewModel.kt** - Manages the news data and loading state for the screen

### Theme Folder

Contains the visual styling of the app:

- **Color.kt** - The color palette used throughout the app
- **Theme.kt** - Sets up light mode, dark mode, and dynamic colors
- **Type.kt** - Typography settings for text styles

---

## 🚀 How to Run the App

### What You Need

1. **Android Studio** installed on your computer
2. An **Android phone** or **emulator** to run the app
3. A free **API key** from NewsAPI.org

### Step-by-Step Setup

1. **Open the Project**
   - Launch Android Studio
   - Click "Open" and select the `first_app` folder

2. **Get Your API Key**
   - Go to [newsapi.org](https://newsapi.org/) in your web browser
   - Create a free account
   - Copy the API key they give you

3. **Add Your API Key**
   - Open the file `NewsViewModel.kt` in Android Studio
   - Find the line that says `val apiKey = "..."`
   - Replace the existing key with your own API key

4. **Run the App**
   - Connect your Android phone via USB (or start an emulator)
   - Click the green "Run" button (▶️) in Android Studio
   - Wait for the app to install and launch

---

## 🔌 About the News API

The app uses **NewsAPI.org**, a free service for developers that provides access to news articles from thousands of sources worldwide.

### What We Request

The app asks for **top headlines** from the **United States**. This gives us the most important and trending news stories at any given moment.

### Free Tier Limitations

The free API key has some limits:
- Up to 100 requests per day
- Only works during development (not for published apps)
- Headlines may be delayed by a few hours

For a production app, you would need to upgrade to a paid plan.

---

## 🎨 App Theming

The app automatically adapts to your preferences:

- **Light Mode** - A bright, clean look with white backgrounds
- **Dark Mode** - An eye-friendly dark interface for nighttime use
- **Dynamic Colors** - On Android 12+, the app colors match your wallpaper

The theme switches automatically based on your phone's system settings.

---

## 📋 Required Permissions

The app only needs one permission:

- **Internet Access** - Required to download news articles from the web

This permission is automatically granted and doesn't require user approval.

---

## 🚧 Ideas for Future Improvements

Here are some features that could be added to make the app even better:

- **Search** - Let users search for news on specific topics
- **Categories** - Filter news by categories like sports, technology, or business
- **Favorites** - Save articles to read later
- **Offline Mode** - Store downloaded articles for reading without internet
- **Pull to Refresh** - Swipe down to get the latest headlines
- **Share** - Share interesting articles with friends
- **Multiple Countries** - Choose news from different countries

---

## 📄 License

This project is open source and free to use for learning purposes.

---

## 👤 Author

Created as a learning project to explore modern Android development with Kotlin and Jetpack Compose.
