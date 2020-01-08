import joel2 as j
import time

place, length = j.starting_url()
row = 1
for i in range(1, 10000):
    for k in range(length):
        place[k].click()
        #time.sleep(3)
        if j.switch_to_about() != 'stop':
            j.switch_to_about()
            #time.sleep(3)
            j.into_excel(row)
            row = row + 1
            print("Trenutno na {}".format(row))
            j.driver.execute_script("window.scrollBy(0, 15000);")
            #time.sleep(3)
        else:
            try:
                ActionChains(j.driver).key_down(Keys.CONTROL).send_keys('w').key_up(Keys.CONTROL).perform()
                j.driver.switch_to.window(j.driver.window_handles[0])
            except Exception as e:
                print("Nije se otvorila stranica")
                try:
                    j.driver.execute_script("window.scrollBy(0, 15000);")
                except Exception as e:
                    pass
    if i > 1:
        place, length = j.other_next()
    else:
        place, length = j.first_next()
