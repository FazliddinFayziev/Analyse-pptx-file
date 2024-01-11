# =======================================================>
# Warning: Values of w or h cannot be bigger than W or H.
# =======================================================>

# ===============================================================================================>
# Left and Top values for SMALLER (2w <= W && 2h <= H) and BIGGER (w >= W/2 && h >= H/2) w and h.
# ===============================================================================================>

# ====>
# Left.
# ====>

def left_left(W, w):
    if w > W:
        print("Error: You are going out of slide!")
        return None

    if (2 * w) <= W:
        return (W - (2 * w)) / 3
    else:
        return (w - (W / 2)) / 3


def center_left(W, w):
    if w > W:
        print("Error: You are going out of slide!")
        return None

    return (W - w) / 2


def right_left(W, w):
    if w > W:
        print("Error: You are going out of slide!")
        return None

    if round(2 * w) <= W:
        value = (W - (2 * w)) / 3
        return (value * 2) + w
    else:
        value = ((w - (W / 2))) / 3
        return value * 2


# ====>
# Top.
# ====>

def top_top(H, h):
    if h > H:
        print("Error: You are out of slide")
        return None

    if (2 * h) < H:
        return (H - (2 * h)) / 3
    else:
        return (h - (H / 2)) / 3


def middle_top(H, h):
    if h > H:
        print("Error: You are out of slide")
        return None

    return (H - h) / 2


def bottom_top(H, h):
    if h > H:
        print("Error: You are out of slide")
        return None

    if round(2 * h) < H:
        value = (H - (2 * h)) / 3
        return (value * 2) + h
    else:
        value = (h - (H / 2)) / 3
        return value * 2


# ============================>
# ALL POSITIONS FOR ONLY TOP.
# ============================>

# =========>
# TOP_LEFT.
# =========>
def top_left(W, H, w, h):
    left = left_left(W, w)
    top = top_top(H, h)
    return [left, top]

# ===========>
# TOP_MIDDLE.
# ===========>
def top_middle(W, H, w, h):
    left = center_left(W, w)
    top = top_top(H, h)
    return [left, top]

# ===========>
# TOP_RIGHT.
# ===========>
def top_right(W, H, w, h):
    left = right_left(W, w)
    top = top_top(H, h)
    return [left, top]


# ==============================>
# ALL POSITIONS FOR ONLY MIDDLE.
# ==============================>

# =========>
# MIDDLE_LEFT.
# =========>
def middle_left(W, H, w, h):
    left = left_left(W, w)
    top = middle_top(H, h)
    return [left, top]

# =====================>
# MIDDLE (CENTER).
# =====================>
def center(W, H, w, h):
    left = center_left(W, w)
    top = middle_top(H, h)
    return [left, top]

# ===========>
# MIDDLE_RIGHT.
# ===========>
def middle_right(W, H, w, h):
    left = right_left(W, w)
    top = middle_top(H, h)
    return [left, top]


# ==============================>
# ALL POSITIONS FOR ONLY BOTTOM.
# ==============================>

# ===========>
# BOTTOM_LEFT.
# ===========>
def bottom_left(W, H, w, h):
    left = left_left(W, w)
    top = bottom_top(H, h)
    return [left, top]

# ==============>
# BOTTOM_MIDDLE.
# ==============>
def bottom_middle(W, H, w, h):
    left = center_left(W, w)
    top = bottom_top(H, h)
    return [left, top]

# ==============>
# BOTTOM_RIGHT.
# ==============>
def bottom_right(W, H, w, h):
    left = right_left(W, w)
    top = bottom_top(H, h)
    return [left, top]
