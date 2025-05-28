# Library
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl

# Inisialisasi WebDriver
driver = webdriver.Chrome()
driver.implicitly_wait(5)  # Tambahan wait implicit

# Inisialisasi workbook Excel
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "buku"
sheet.append(["gambar","Author","Judul_Buku","Harga","Deskripsi","Penerbit","TanggalTerbit","Kode_Buku","Halaman_Buku","Bahasa","Panjang_Buku","Lebar_Buku","Berat_Buku","Link"])  # Header kolom

try:
    url = "https://www.gramedia.com/categories/buku/komputer-teknologi/manajemen-database"
    print(f"Membuka halaman: {url}")
    driver.get(url)

    time.sleep(2)  # Tunggu manual untuk load awal

    # Coba tekan tombol filter jika ada
    try:
        filter_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="productListFilterBubble#1"]'))
        )
        filter_button.click()
        print("✅ Tombol filter ditemukan dan diklik.")
        time.sleep(2)  # Beri jeda setelah klik
    except:
        print("⚠️ Tombol filter tidak ditemukan, lanjut ke proses berikutnya.")

    # Coba tekan tombol "Load More" jika ada
    try:
        load_more_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="productListLoadMore"]'))
        )
        load_more_button.click()
        print("✅ Tombol 'Load More' ditemukan dan diklik.")
        time.sleep(3)  # Beri waktu untuk load data tambahan
    except:
        print("⚠️ Tombol 'Load More' tidak ditemukan, lanjut ke proses berikutnya.")


    elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located(
            (By.XPATH, '//*[@data-testid="productCardContent"]')
        )
    )
    print("Jumlah elemen ditemukan:", len(elements))

    for i in range(len(elements)):
        try:
            elements = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, '//*[@data-testid="productCardContent"]')
                )
            )

            # Scroll elemen ke tampilan
            driver.execute_script("arguments[0].scrollIntoView(true);", elements[i])
            time.sleep(1)

            # Tutup chatbot jika muncul
            try:
                chatbot = driver.find_element(By.CSS_SELECTOR, 'img[alt="Gramedia"]')
                driver.execute_script("arguments[0].style.display = 'none';", chatbot)
            except:
                pass

            elements[i].click()
            print(f"[{i+1}] Elemen berhasil diklik.")

            gambar = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@data-testid="productDetailImage#0"]')
                )
            ).get_attribute("src")
            print(f"Gambar: {gambar}")

            # Ambil data dari halaman detail
            author = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@data-testid="productDetailAuthor"]')
                )
            ).text
            print(f"Author: {author}")

            judul_buku = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@data-testid="productDetailTitle"]')
                )
            ).text
            print(f"judul_buku: {judul_buku}")

            # harga = WebDriverWait(driver, 60).until(
            #     EC.presence_of_element_located(
            #         (By.XPATH, '//*[@data-testid="productDetailFinalPrice"]')
            #     )
            # ).text
            # print(f"Harga: {harga}")
            

            # Ambil harga final
            harga = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, '[data-testid="productDetailFinalPrice"]')
                )
            ).text

            # Ambil diskon (jika ada)
            try:
                diskon = driver.find_element(By.CSS_SELECTOR, '[data-testid="productDetailDiscount"]').text
            except:
                diskon = "-"

            # Ambil harga asli sebelum diskon (jika ada)
            try:
                harga_asli = driver.find_element(By.CSS_SELECTOR, '[data-testid="productDetailSlicePrice"]').text
            except:
                harga_asli = "-"

            # Cetak hasil
            print(f"Harga Final : {harga}")
            print(f"Diskon      : {diskon}")
            print(f"Harga Asli  : {harga_asli}")

            deskripsi = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@data-testid="productDetailDescriptionContainer"]')
                )
            ).text
            print(f"Deskripsi: {deskripsi}")

            Penerbit = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@data-testid="productDetailSpecificationItem#0"]//*[@data-testid="productDetailSpecificationItemValue"]')
                )
            ).text
            print(f"Penerbit: {Penerbit}")  

            Tanggal_Terbit = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@data-testid="productDetailSpecificationItem#1"]//*[@data-testid="productDetailSpecificationItemValue"]')
                )
            ).text
            print(f"Tanggal Terbit: {Tanggal_Terbit}")  

            kodebuku = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@data-testid="productDetailSpecificationItem#2"]//*[@data-testid="productDetailSpecificationItemValue"]')
                )
            ).text
            print(f"Kode Buku: {kodebuku}")  

            Halaman = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@data-testid="productDetailSpecificationItem#3"]//*[@data-testid="productDetailSpecificationItemValue"]')
                )
            ).text
            print(f"Halaman Buku: {Halaman}")  

            Bahasa = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@data-testid="productDetailSpecificationItem#4"]//*[@data-testid="productDetailSpecificationItemValue"]')
                )
            ).text
            print(f"Bahasa Buku: {Bahasa}")  

            Panjang = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@data-testid="productDetailSpecificationItem#5"]//*[@data-testid="productDetailSpecificationItemValue"]')
                )
            ).text
            print(f"Panjang Buku: {Panjang}")  

            Lebar = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@data-testid="productDetailSpecificationItem#6"]//*[@data-testid="productDetailSpecificationItemValue"]')
                )
            ).text
            print(f"Lebar Buku: {Lebar}")  

            Berat = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@data-testid="productDetailSpecificationItem#7"]//*[@data-testid="productDetailSpecificationItemValue"]')
                )
            ).text
            print(f"Berat Buku: {Berat}")  


            link = driver.current_url
            print(f"Link  : {link}")
            harga_excel = harga if harga != "" else harga_asli


            # Simpan ke file Excel
            sheet.append([gambar,author,judul_buku,harga_excel,deskripsi,Penerbit,Tanggal_Terbit,kodebuku,Halaman,Bahasa,Panjang,Lebar,Berat,link])

            driver.back()

            try:
                load_more_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="productListLoadMore"]'))
                )
                load_more_button.click()
                print("✅ Tombol 'Load More' ditemukan dan diklik.")
                time.sleep(3)  # Beri waktu untuk load data tambahan
            except:
                print("⚠️ Tombol 'Load More' tidak ditemukan, lanjut ke proses berikutnya.")

            time.sleep(3)  # Tunggu halaman utama termuat kembali
        except Exception as e:
            print(f"❌ Terjadi kesalahan saat memproses elemen ke-{i+1}: {e}")
except Exception as e:
    print(f"❌ Terjadi kesalahan utama: {e}")
finally:
    # Simpan file Excel setelah semua halaman selesai diproses
    workbook.save("data_buku_gramedia.xlsx")
    print("✅ Data berhasil disimpan ke file Excel.")
    driver.quit()
