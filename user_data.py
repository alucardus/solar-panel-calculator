# Creating User class for each user object
class User:
    def __init__(
        self,
        address,
        city,
        state,
        country,
        rooftop_size,
        annual_consumption,
        module_type,
        array_type,
        azimuth,
    ):
        self._address = address
        self._city = city
        self._state = state
        self._country = country
        self._rooftop_size = rooftop_size
        self._annual_consumption = annual_consumption
        self._module_type = module_type
        self._array_type = array_type
        self._azimuth = azimuth
        self._annual_output = 0
        self._price = 0
        self._monthly_output = []
        self.modules = {"0": "Standard", "1": "Premium", "2": "Thin film"}
        self.arrays = {
            "0": "Fixed(open track)",
            "1": "Fixed(roof mount)",
            "2": "1-Axis",
            "3": "1-Axis Backtracking",
            "4": "2-Axis",
        }

    @property
    def address(self):
        return self._address

    @property
    def city(self):
        return self._city

    @property
    def state(self):
        return self._state

    @property
    def country(self):
        return self._country

    @property
    def rooftop_size(self):
        return self._rooftop_size

    @property
    def annual_consumption(self):
        return self._annual_consumption

    @property
    def module_type(self):
        return self._module_type

    @property
    def array_type(self):
        return self._array_type

    @property
    def azimuth(self):
        return self._azimuth

    @property
    def annual_output(self):
        return self._annual_output

    @annual_output.setter
    def annual_output(self, annual_output):
        self._annual_output = annual_output

    @property
    def price(self):
        return self._price

    @price.setter
    def price(self, price):
        self._price = price

    @property
    def monthly_output(self):
        return self._monthly_output

    @monthly_output.setter
    def monthly_output(self, monthly_output):
        self._monthly_output = monthly_output

    def get_module_key(self):
        for k, v in self.modules.items():
            if v == self.module_type:
                return k

    def get_array_key(self):
        for k, v in self.arrays.items():
            if v == self.array_type:
                return k
