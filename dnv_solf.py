#DNV, Stdandard Offshore Leak Frequencies
import numpy as np
def F_full_pressure (d,D):
    f = 5.0E-5 * (1. + 3000.*np.power(D,-1.9))/d + 3.0E-6
    return f
F10mm_6in= F_full_pressure(10,6*25)
F3mm_6in= F_full_pressure(3,6*25)
F3mm_6in - F10mm_6in

# F10mm_6in= F_full_pressure(10,6*25.4)
# F50mm_6in= F_full_pressure(50,6*25.4)
# F10mm_6in - F50mm_6in

