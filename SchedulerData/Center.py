class Center:
    centers = []

    def __init__(self, name = 'Mathnasium of Stafford', street='263 Garrisionville Road (Ste 104)', city = 'Stafford', state='Virginia',
                 zip_code = '22554', email_address='stafford@mathnasium.com', time_zone='America/New_York'):
        name = name
        street = street
        city = city
        state = state
        zip_code = zip_code
        email_address = email_address
        time_zone = time_zone
        mailing_address = name + ', ' + street + ', ' + city + ', ' + state + ', ' + zip_code
        self.centers.append(self)
