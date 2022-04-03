from src.utils import color as ucolor


def test_color():
    c = ucolor.rgb(133, 144, 254)
    assert len(c) == 3
    assert c.isvalid()
    c_int = 8753406
    assert c.to_int() == c_int
    assert int(c.to_hex(), 16) == c_int
    c2 = ucolor.rgb.from_int(c_int)
    assert c2.red == c.red
    assert c2.blue == c.blue
    assert c2.green == c.green
    assert c2[0] == c[0]
    assert c2[1] == c[1]
    assert c2[2] == c[2]


def test_rgb_hsl():
    c = ucolor.rgb(133, 144, 128)
    h = ucolor.rgb_to_hsl(c)
    # assert h.hue == 0.28124999999999983
    c2 = ucolor.hsl_to_rgb(h)
    assert c2 == c


def test_rgb_hsv():
    c = ucolor.rgb(133, 144, 128)
    h = ucolor.rgb_to_hsv(c)
    c2 = ucolor.hsv_to_rgb(h)
    assert c2 == c


def test_hsv_hsl():
    c = ucolor.rgb(133, 144, 128)
    c_hsv = ucolor.rgb_to_hsv(c)
    c_hsl = ucolor.hsv_to_hsl(c_hsv)
    c2_hsv = ucolor.hsl_to_hsv(c_hsl)
    assert c2_hsv == c_hsv
    c2 = ucolor.hsv_to_rgb(c2_hsv)
    assert c == c2
    c3 = ucolor.hsl_to_rgb(c_hsl)
    assert c3 == c


def test_lighten():
    # https://mdigi.tools/lighten-color
    color = ucolor.rgb.from_hex("eeeeee")
    b_color = ucolor.lighten(color, 5)
    s_hex = b_color.to_hex()
    assert s_hex == 'efefef'

    b_color = ucolor.lighten(color, 63)
    s_hex = b_color.to_hex()
    assert s_hex == 'f9f9f9'
    
    color = ucolor.rgb.from_hex("24c1cc")
    b_color = ucolor.lighten(color, 45)
    s_hex = b_color.to_hex()
    assert s_hex == '81e2e9'

def test_darken():
    color = ucolor.rgb.from_hex("eeeeee")
    b_color = ucolor.darken(color, 15)
    s_hex = b_color.to_hex()
    assert s_hex == 'cacaca'
    
    color = ucolor.rgb.from_hex("9b3c08")
    b_color = ucolor.darken(color, 95)
    s_hex = b_color.to_hex()
    assert s_hex == '080300'

def test_is_dark():
    color = ucolor.rgb.from_int(16777215)
    assert color.red == 255
    assert color.green == 255
    assert color.blue == 255
    bright = color.get_brightness()
    assert bright > 128
    assert color.is_dark() is False
    assert color.is_light() is True