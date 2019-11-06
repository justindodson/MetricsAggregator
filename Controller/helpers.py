
def region_switch(region):
    regions = {
        1: '1 US',
        2: '2 US',
        3: '3 US',
        4: '4 US',
        5: '5 US',
        6: 'CA'
    }
    print(region.get(region, 'Invalid Region'))
    return region.get(region, 'Invalid Region')