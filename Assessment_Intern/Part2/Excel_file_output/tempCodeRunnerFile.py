try:

        descli=driver.find_elements(by=By.XPATH, value='.//div[@id="feature-bullets"]/ul/li')
        desc=''
        for item in descli:
            txt=item.find_element_by_tag_name('span')
            desc+=txt.text+' '
        sheet['A'+str(r)].value=desc
    except NoSuchElementException:
        sheet['A'+str(r)].value='NOT FOUND'
        pass
    try:
        asin=driver.find_element_by_id('olpLinkWidget_feature_div')
        sheet['B'+str(r)].value=asin.get_attribute('data-csa-c-asin')
    except NoSuchElementException:
        sheet['B'+str(r)].value='NOT FOUND'
        pass
    try:
        pdesc=driver.find_element(by=By.XPATH, value='.//div[@id="productDescription"]/p/span')
        sheet['C'+str(r)].value=pdesc.text
    except NoSuchElementException:
        sheet['C'+str(r)].value='NOT FOUND'
        pass