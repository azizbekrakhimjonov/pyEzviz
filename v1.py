import pygetwindow as gw
import time
import datetime

# Hisoblagich
youtube_count = 0

# Oldingi holat (takror sanamaslik uchun)
previous_state = False

print("YouTube monitoring boshlandi...")

while True:
    try:
        active_window = gw.getActiveWindow()
        if active_window:
            title = active_window.title

            # YouTube ochilganini tekshirish
            if "YouTube" in title and not previous_state:
                youtube_count += 1
                previous_state = True

                # Faylga yozish
                with open("youtube_count.txt", "a", encoding="utf-8") as f:
                    f.write(f"{datetime.datetime.now()} — YouTube ochildi ({youtube_count}-marta)\n")

                print(f"YouTube ochildi! ({youtube_count}-marta)")

            # Agar YouTube yopilsa holatni tiklash
            if "YouTube" not in title:
                previous_state = False

    except:
        pass

    time.sleep(1)  # CPU'ni tejash
