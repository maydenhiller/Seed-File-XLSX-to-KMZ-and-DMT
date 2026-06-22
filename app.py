# app.py
# Seed File XLSX to KMZ and DMT
# Upload a Google Earth Seed File (.xlsx); download BOTH a KMZ (Earthpoint-style)
# and a DeLorme Street Atlas .dmt with AGMs, Access, Centerline and Notes layers.
#
# This single file merges two apps:
#   1) XLSX -> KMZ generator  (builds the KMZ in memory)
#   2) KMZ  -> DMT generator  (consumes that KMZ and writes the .dmt)

import os
import re
import io
import math
import time
import base64
import struct
import zipfile
import xml.etree.ElementTree as ET

import streamlit as st
import pandas as pd
import simplekml


# ===========================================================================
# PART 1 : XLSX -> KMZ  (library code from the XLSX-To-KMZ Generator)
# ===========================================================================
# app.py



# -------------------------
# Gibson Integrity logo (embedded so every KMZ includes it automatically)
# -------------------------
GIBSON_LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAfQAAACyCAYAAAC0oD1PAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAI9hJREFUeNrsnc9vI8l1x0vcuS/3L9gWjJyHQi5GjGCaAdaIEQdD5egchgySIMjBIx6CbOxgRAlZYC8LSUAW8Y8g5PgQn2JpYHvtwEDUcwkMxLE4cGAEDgxx/oLh3kfD1JNez7Q4TbJ/vOqq6v5+AGI0EkV1V1e9b71Xr14pBQAAAAAAAAAAAAAAAF7wf7/9bYBWAC7TQhMAAEAmOlrUO2gGAAAAwH8vfaRfbbQEgIcOAAB+c6xfR2gGAAAAwH8vvfO/v/nNLVH/m3/8YQ8tA+ChAwCAR/zOF74wvbq6+vx/fv3rfvy9rS0V/e2nPzzSL4TjgTW20AQAAJCfZ7/61cWrV68GO3fvTun/f/dPP2ovlNpTC3X28V9/dYoWAgAAADzgvy8uwl/88pcX//WLX9zyyr/5rR+Pv/GtH/fRQqBqEHIHAIAC/O7OTnT18iWF32+tp3/0V380eGdrq/333/5sjFYCVYKQOwAAFOQ/f/7z9qtXry4Xi8Xw97/0pUnyZ/vf/Ul/sVAP9Zfdw7/8yhytBeChAwCAo/zeF7841176iX4d/UcU3So6c/AXX5ncaW2dvNPauhh99ycoSAMg6AAA4DJXr14dv7y6ml9dXb0VYn/053+oRb118E6rdX74zz/F1jYAQQcAAFf5gzAkL/1Avzo/+elP3yo6880/+/LkzjutoX6d/sO//PseWgyYAmvoAAAgwI8+++xSLRbBYrHo/vFXvxot//zjyc/6C6XG+ueTbwy+PECLAXjoAADgINpDf/zy5UulX+N/+8EP3iow82H/g4n2oIatra2+FvdT/UIRGgBBBwAA5wT96uqYBF3/G+jXftp7tKhTLfiJftF6+jlEHUDQAQDAMf5kd5cS4yYs6nv/+v3vhytEfcCi3mFRRwY8gKADAIBLvOSw+9XNa/z4e99b5YEP9WsKUQcQdAAAcJA//drXqHrc7OXVlXq5PvROhWa6+kX/tiHqAIIOAACueelXV2cJL33v29/5TmeDqKuEqGOvOoCgAwCAE4L+8uVTznYncVfLtd6XRJ3C7sOEqFP2ex+tCIqAfegAACDMp59+utCoBf2H/l0sBl//+tcnq95P29jUTeZ7zIC2uaElATx0AACw66VH1945e+n6dfTJJ5+s26JGme+zxP/H8NQBBB0AACxzdXU1vYrD7jdr6W3ayrbq/byevlw9DqIOIOgAAGDZQ3/+eg39jbA//Oijj4I1oh7pf44h6gCCDgAA7gj6NM50f+2lX121V21jS3CgbofeIeoAgg4AABYFnQQ8Kebx1/1Hjx6t89LTQu8QdQBBBwAAG7x69WoaC7n++vpFWe/MWi+dQ+9nKT86QvEZAEEHAIAKGY1G8zU/XuulM8OU76GiHICgAwCAY2zy0mfq7QQ5iDqAoAMAQJVoDzzc8Jaefs+mo1MpQW6+QtTHOHoVQNABAMA+JMb9DV46ifnJih/Hp7RB1AEEHQAADBJkeM/DDO85XuGlx6I+RlMDCDoAANgV9GBTaJ699LM1b+lpLx2iDiDoAABgiHsZ3/cgw3sONvy8jz3qAIIOAABmCDO+b+P555zxHm142xiZ7wCCDgAAgjx69KiX4+3tjO9/nOE9lCQX4AlA0AEAAMhwX/r9fDb6fNPkQL9O0fwQdAAAAOW9cxLVXs5fCzO+7yzDezpIkoOgAwAAKE+PPeU8BBlKwRJPMn4ekuQg6AAAAEqyX/D3NnrpH/Y/OMvxeTjIBYIOAACgCNrLJq84KPjrdzO+L6uoozwsBB0AAEDF3jmR1Zt+lvMz9/FYIOgAAACye+d7JbzzPIIe5fzcPe2l9/CEIOgAAAA2i3lbwBPOFBr/sP9BVOCzEXqHoAMAAMgimCp/ZvuqiUEWpgUmC9jKBkEHAACwRoQpnC0V0s4adp8V+OweQu8QdAAAAOliHljyfJ8V/D2E3mvOHTQBAGCZxWJB3mJs/MOUt0RL/59vbW1NG9ZMp0og1F6AWcHfo2s90q8BejgE3ZZhoVlwkDAqyWMJwxwfRcZmnvj6cx4Y1y9tjGboDqCh4h2PLxpbHZUt9Luf8jmvxV29Wed9mjL+VuHNpEB750cqe4jcFUEnqIrc44LJdUX61jm3U7eBE75mC7p++O0lwxIKfnxn3USAjVHEg+U5fz3VnXDuWBv1VbntMUYMjG6nSYZrD4WfqSS3DJy+n6jOA5/HGq2pPjQgTO3Ecw5zPoOuB2JOY3DP48dPywTbFf2tp9wHznWfG+hxdVaDsRM4aIOvbdYdBxqHjMkDfui2SxWGyx6Ivr4ZGxrqmJEDnvwDB0XxOEf7ulroYn+Ftxl7lk/566nP0RwW8j0WctfWU5334LSYd5ShdfPDw8Osk8iy/S/QXvpIe+mjCppsnpjknbKoT5Tf9B21Y1tbFmc4D9lDCDx7mLHAP2GBn1tsQ3rd5w7WrvD+T9j4Fl6qYG89cHSCkrkP+ORx6DbfY0PkamLUgW7Pkavtx0lwF6baTwt6ZnusBXkhILTbWtTnhvvctXe+9G2vRT3FQw/ZDlflkJLtecz299Yk8E7FDdH31IDfmt2ygPb5ns5Y3M+qFHcW0mth0ddwoG6SXfq+GNxER5zo6+8pob28VfcBfe30zKkPnLi6RsgGaOz5uLMt5vFZ46b6aNV9p6oEubT7GlMEzFdRT9jepMCOKrBjc54MrXQitioyKHGIwpQ3HvHNrtvOESfTdQw1uHXDrtt5bFDUjc6qeenl3CNRX9UPD1xafxcyMlO+t+cJA32dX7LkrdAzfF8VXz7rupi7wGJ+btgDm2gPPbO4CnjoMeSlzwz3wYUNm2JpvMVLMtJ9Za4yJBbeMXxzJoR8pkquafM6Ypx0d0/Ic2knvDa6vmHVwq7/3iCRsSztmU8MX/uU1tfYC/IVaveQn//A9lo7j79xCQNCSyuTdfex5K1ES1GBOOkuUH5TRUb7M0v3Rv3DdCLidEX7kaceuLzMUsKOSTsnmfRky5AhCYUHgXHvlz2Z+4IerpX1QDakl5ITKH0f2xVe/7nghIT6yjDl+3GU5i7/LVMRG+oDxzYMS0kxpyWcY6klpBwT+x3Xli20d24y6nXr3rWHnune+ZzzC8G/3TW5jS3DmKZJY632xut7Him5xDlyXDNNulrCN9HWryMlF56as3HZpgducrDTugR3qvf4b8597EjsMUkmalWd9HUiKaoUwk15kViN9GtXv+h575BREb6P6zVKXgap2piEBcV8zqI6kswH4ejOxjZusJjPsop5om9JYjtju29jnBhGciL/OOsbW8JGhGaNUvszJyzko4qTzebsWW97LOyS4bsnFV97VHVjkZDwZG7bwASmUmPFy0lFli0yrdGVHFeDFRET59BiPqpIzG1MmpcJtdcfGvz8qWvjpAodEbRlmT+nJWRERuyVBxKzVTYsA5tFXZaEfeJZf5p7PhCsRTfIa1fymb9VGquiCXCVVPLiJYjjFePeFTHvV+y1Ps75fhPr+Sbv9/MGe+oidqkSQecQ+7lgZ6CZ6o5Lma4Jz2LXI6FEicVyz3xiSNSNVhfjKFmRE7UqTeDUf2uY4nU4Iegs5lWKSt5wu1Jmcj5CXpu3DY2TU440gZy0ShiPeJtRKHQtE17TdFI0ee/fDsSyUaJ+IO0FcdKiKYpMrKeWEvecS4KyIOaqYB+7a+haHjryKGhSeg5Rr0jQE2IuNaMb+JDlyKGPLkS9MaI+En7WbWUotMljssjk+sTiWJokJxaWxbxnQczj3Tt5MTUppINbAkeGXweiXoGgGygA4lVxAY4gQNSbg3QSV9+Ql/6wYH+2OfaSk4nPbV2Eyfrsm+7/8PCwSETSZGi879DYg6ibFHQDYj7xsVIQRL1RXnqk5Nd3ewYutchnRpbbdqosr52zmNuoUEg2JPdSh+Fs9MITQ4i6Z4Ke2A4j1bBnPhcTYFEfKI8zykEhT1KCB5IfxhPtIuNy5kDbTm1NLviwFVvlhg8KeuemBb2tJw09x8ZfLOoBTJGchy61LS02JN5XBmIPY4BuVHukxaYj7HEUNfLPHWhbKyVPKzhsZa3902JeNBHxbgXX98DBMXhdHY8nr6CMoPPeQMmGdDabvYCoU1LLGbpSfeGJm3R/lRxPRY38PYeauWp7YPqwlXWUcQKq8J572kt3McTdZk8dol5U0LnGeV/wbx64esxkyQGK0Hu9ke6zoeBnBRX/nuRkiSpBblVpE7ikqy1RONPeeVTkFysOhfcdHYcQ9aKCnjhHWcwo1ulknYRRik+mAhB0GxQWdC5G0xgqLumaFoUo453fr/BaHzj8GCHqBT106cPahzVuy2N46bVGeluVZLi7jKe935QHyHvNbd5v0US4mConXx2H9qSvE/VGTUgLCzqH2iUba+JSSVd46QDciITpkrSOiHmg7Ow1j4lKJMLF4faqBdZ1sYxFvY9hvNlDP5KenTagPY/RpWrLrMb3dtSA8KWtjHaibKiduG/huu978mzHEPU1gs6NIzkbnOQ5McZzL32CbgVB93CCUNs1Se2dk3Ni894o1F74eXPGuQ3B6nn0mCHqazx06XWmJoWinyhQR6QF4blj91fL8CWvm9tcUigVamesPZMKKtNB1E0KugHvfFrDbWrrvHTak47kuPohHa6VHBMzwXsko3hUh1KbXDzG5rq5RKidsFmONfTssY+bkBOSx0OX3q7QxESxSFVT0Qn4i4uCHkMG8aIGGcRHyt66+bXtKxNqZw/ZRjJckns+PncuhtZsQed959KDuHFV1Phc911oVq2QnKDNhHNKTETAyBZQCP7UR29de+dkx/o2J2xazEeee+eEr3kV/SaKestw54nqUuIVNB5JL0l6kvvU4H2Th3jJ21h9885tUrrmBq9fh5bvgw5rgah7KujSgxYJYsB72EOVNGrSy1CRaaOuX6e+eOvaO+9b9iwnRcu7LuFK0Z/A4+HbKFFvJYxWx8CDixQA/iM50RXfwslRsLOK2uHSg6Qjm0I4F/LOpQt7lcEVD512CxRZXmqMqCc9dOnOM29SdjuoNZIFNkwVWKoqGkYeOiUdOXk+NXvnNq/rpGR515gjh5rVlQRfKr/cLegokqhf1GH3RlZBl85mhJgD72HRkvLQj00VWNKfO1HVFqshB+DCQW/dtndeulqk9s73lFthbmdEkKJR+kWiPikYaTivs6ib9NCfKgD8R8pTmirz5Y+rLq+c9Nath2W5iIzX3jlXhXPtwJzQtUGpRX0AUV8h6OyFuFw4AwAb3nlfyDu/LjBieseHBS992VsfWX5kNo/8FPHOlfwJl7UFor7aQ+8Y6uAA+CrmHSHvnMZBt8J8koHFZtvndcrKvXU+Tc3m1rqJgHceKkfrp3PkAKLeVEGv81GpoBGe+bmApzStWMzjcWfzZMOOJW/dthCW2orIgulyJraze9FZ1I8L3tNlnQ4ligX9fZhxACFftKmOuZIJex5XLeYJAzdS9k/9I2/9ssLysTaP+pyWLfGqbtbNA4zCwn2etgoWiU7FhxLVQtRjQZfuSFg/Bz565Req3Mlc8fG522RgbFZJLBGKlCRgY2n0sBc+hCW0eJ+PS3rnobJ7IlxdRH3SdFFvGfpcrJ8DH7zxHovNJXvlRSe2VNSF6ve/R0Jqamuap6KulPnDXkLL91e2qM8YIxKiLinoIboCaIIXzlusFvq/L/TrlMUmKPnRtH57yp89cskosKgPHbiU2FsfG/DWbbb3rEy4XXvnI4VQO0TdcQ8dABchw0v1EQ7Yq4qUbDSJJsb77I1esrhbz6LVBo7W83eUG0thfSV/2IvNIz7LtukDDEujop53fLd5/PYh6HKdHAATg5xO/xvxi0LklLT2nv7RNg/+iaDAByzuL0yvIWe896l+7fBkxvaSWPKwl0CorW3xrIR3brsQThNEvVuwv499FHVTgv45uhPwaODT+eQTDk/H4j4T/BN7ypEjSDkDfke5cXBST8mUj7UpimXa0SfvfOrp2J42SdQRcgfgtgGYs7hvK9n93LFXOnbAW59xPeyhI956XD42d7s8evTI5ySm0JcL/bD/wdzjMd0YUb8DE/669K0rDy1CUR5nDAGtgdNaO2UhSwkH9bOO/tyuzW1tfH/HifuzLS7090nUBzn37tvOUSjkuX48+VlH+VPidV6DsTylMaeKFYwiUY9D+BB0DyBBd+lABAi6m4ZAStTjamq7to8Y5i12XQ5771sWmbgcZ9eXo5dLlHv1KbIwreFYLiLqd7mAjbPEIfdZw402ecWvUTfrqEPD7UKGgEK6O1u3GUFGnesf1/XYhQ0bTSJPXaklnciEtz2ZbLPxrPsBJQEE3Y6os30vck97tGTWREF/3/OHPksYuImhAbLN2dbYEeCPqO8q2fBjoBw6ICKxtj5QdsOs5L2e1rxLvevRta7N5Pctj6HkBL3vsqibSooL6jDi+MFLe+rx6VuopufhRE/Jn2ZGxnDfsfucsBdzZvEyQoHsd5epU8g98HAs11LUY0F/CnO99sGfCH7kBGLudX+IC9JIslfhISaZ+z3t1TcQlcjDvuuhd+2dhjXv8rMP+x9sEr25p2O5dqLeMvRA6rb+JRkWxx59/zFxPOmRizfKExhb3nrb1XZxzKaYJMvkNfD1IQiIulNnqrcMda6OAqC+XnpkwEvvuLrf1bK33ttgMG17h0VtnS8T+yd1d+C4fxfNlwqVQ3kwpgRdNSBLFTSbxwY+0+nKYZa8dbIjKyvsHR4e2vZ0iyYAzzzo4xRuz/Ks67KtrejphB1XRL2VCDtIdzB46aDORAY+MxSqbV43b/3+hp/bzsgvgg8imHXSGtZlUPsu6i2DBgqCDmoLZ7yb8LJ6ntx/ld76pjaxKY6FxCxDopltaJJ03NCx7a2oJwX9ifBn34XZBzXHhFG+78vNL3nrM5N/a8MuAKviWGIf9pnDj/fE5/rtlkX9wtaZ6iY99FABUG+eGfhM7yJb7K3HR7OaIqj4OVRh6544+kjzeuf36ji4WdQHBfvqubKQ/d9KzraFZ4yB6+uBAAgYPmnaPo4b9tZHLOwmPObAVQ9dFY+qnCk393AfZPXOHz16VGsbz0WWioh626qgG5oxwksHdcaUkHhrJKmUcYktQIXgTHebwhhqYcu9bsqi6VrYfaqvK493HqoanMZmSNQrp5Vy4ZIP574CABiH1pi5yIUTIfsS4cpVbKp9blsYiyYzHjjWlfI+Mwq3P6v7+PJF1NNquUvOrHvYjw5Afo+vwO/s8+85M96EjWCw4ee2y1c/LPJL2hueKXeyyQ8KZN/36u6h+yTqaYJ+Ivw3+rDPAJj1zpWjy1tsBCUEa5MXaNtD75TIdj9wQBQjLeajPL+g77fHE8jGnBjJ/bnr6iSmlXLBM2Ev/aECADjnHVZoBIemjf7h4aEL69FFvfS5Zc8vPhq46P026ghoLv3spKi31swYpQhcO0kKACFM9evMBpIz4n0oRjOs4G88tnyP/aJZ31xi1Ubo/fpwkrx7zvk+qf/PeDLVKCj500VRb6242JmwqO8rAEAeI5uVZTEPXLwhgQNtZhm8dBe2gRW2dVpUhxaiDLsFq9bFp+BFTR2kLop6a83PjpVc9acQXjqoIaYKauQxEMs7SVyu0FjGg85qi04s32O/xFo6QaH3KkLYcxbz3ILMZ8DHE8mnTTYArol6a82F0gVKhsngpYO6YcQbZiORebK84f8uMa3gd48dMK7jEl56fD73xLCYdzOepLYs5u2l+4uabgR4vJoqqCTmocclHaU6Vujqec8A5IXXrk0IepTjGtLEu+PqVtGcE5Vb3jk7GBtxJDmOMt5HZURdv8hTN7FH/Vp8ShwOc5To91Pd3jNYg9fL1F3bot7K8B7JDNUj7EsHNcFUIlqeEGZQ8bXZIq8X6MI2sP2SoXfF28ikPD9qD9pnvsN733Oj74ccsqRT9hhm4Jaoz22LeivjRQ6EBgiJ+SkePagBDwx9bh7vcpWg122raK6S1Ow1njhw3edFSsIuiTqVYt1hG1xEiMluT/RrO+8+8yUxp8nJ8lLCBGbALVFvZbzIqZLbJ0mh9z3PnlOArgpiuLyqiRKrUc6w9KoEuI6LSagFD52Z89JfXiSTess4MKVFnYV9ol/bLBbHGwQjXnYYsJAPyhyFymJ+vizmTdyu5rqo38lxkWd6QA5UiYSPBBR6n/JWliZ7Y8BPTCV45g1hrhMKWuvccazdikyCCq2Hk9hoIaLlwlMH7plEvSshgJyV/tpufjz5WWepH0wlzzFPiHm7ZF9tnKhrjevyOOxX9ncLzLL7QqJ+PYspkShTlVcRpsxOy3DAx0z6cs3dqide+voXgh5v1/H+8NoQ8yllea6FrmOdJz7Un3ns0FgaFzBu25xwVFSQSNBdyCm43t7kk1fLa+ZHKWIe6fvo5njuI6FJsHHb6VC/T04OMut0q8CHT5RM+P06HOXK6VArHkRbaPICaoDh/mCiktq+K+OL2y6vsE7KiDkjlf8j5akHnoj5Eff1tCjQAaxBLs2kPjip4m+1Cl5gU0SdZpXSA/B9dHFvOVVm8ikODEVBricgjuws6at8J8HNJYSDPWJXTsgiO3fBh5q4KuS05e5Cf7kqz+lMt2kEU+CmqLdKXKC0qIeOeWP9NZ26DAG6t3+eOYfNTPTRyHAYMS2hyYZ3njfkeiDgnceibqtO+ip7d0oesESynKCQt3nv/IVan+swhEVwV9RbJS+QLm5HlQ9pxaLuRPa7YJ6ALUHHXn+5vhCwIPYNfDytq+5WcBuU9X5q0VM/zdknz6TX/rWoGz/xLSdk6y55ndq2mPdZyDdNug5QSEZE1HNFnvKM2zsCFzjVf3CbB21ZD4ay36k29UBqdl7AgB8Z8sxfCzo9oKyVr0p4ZaB8X9hjI2dCCK+TpEr2g2mOMdfjvtc13PeW2zBvZENyi+wyXRauwJEudr0kogV1n8VyUtUf5ugACfnDjO1BVeFGsAoioj7S42KWw2kkex4Z99ATFzjnbGKJZAka/BecGVml8ab98ReGxbwqwX3fo2tdfg5WJyMcXt/Tr0uVnuErASV77QgI67MCz/JS31uvona8yBnZkJjkrPPS43O/Xcs0D1jYLzkUb2zCQev3+kVC8kLdLuO6jrJ5CO8KXf67qiYILlnf/lxDBnksJAQzniScmRrkHFLdVxXuFVQGt19weOZSUIjEt35tuP6RktvnvfHaub+2eSJ5T5k93OTaMBYslLKq714WbRtlKBLGE4Zxzj5oVMyXRG3V3mqXmPIzesLe8bzEvSb7dpF7HpSJHvDkWGKSQjX9t+vkrWdc3s28dXjL4IVKhivjykcnUvvW2ehQwRgbGafGRFJYEHN3KMcmIy5xzBO5uQWDsEnYTyQmGTymHhaYFJFYDKtcCvBE1JedG3o9TdjE6ZKXH4vm3cQktXS/5fwDW/3Tii1yTNQH7NHbE/SEgd7jQS41cOZshJ5yh55uMgR8HXG5zjIzVUneM2Dcr7fFGPIsjRcBosQtVb+DRSZKMGN7zXMvu6Xu1rjKYjQTJXDv8XNrF/ibYhGLBoh65X1Xi/mgRL80MUGncbRT5eSvIlFfF9XKHNHdquhiTQj7qhlskljIXWRX0pCxcTVpnIyKetlqSo4RH4hxUlVyZ2JrmHQOyFS9vebcEehnlXvlEPXcz71wZTvuj+eG7G9lyzMVi/qqvjjh7Hg3BH3pIcdhOd8ysckAUf3i+0JGk/ICdoXaNF7eqIJjSaHi+gNHyv/M/HhZ6Iktj5PbM1DV54Q4FbEoIOoBRziwO0RGzPvKXEJpcrwNs4aiPRf1zEu0W5YvPF7DDhxt32sDrRJJecJh7dxhdx4s1F7vsgEKLbVNxAP/c+5wUUbxjq/3nrq97ucbM77/pyr/KWlVTZ77PMZsC1XlEYsCoh6X9a3bkk+Rcb2bRcx58hhPHGN7JBG9ydu34nGY9Ghnvj6AFFGnXWTvOS3oKTcQKvvr23Fm6dN1XhZvx5EwkrkPz8hwIIeViU+WaIOhhD2TRiIp3s/Vm2WdqU/hPja8vYrH14zHktWIRQFh96WPGome5FkzN3hQUVm8T5xLEfVMzt+WwzdDr7uGZn2xwSaj80xlTAJa8pIlMjdrtw0DeCPwyTEWlJygxuNpyuMp8tlD4nV1U3X7XWWoxfwYo8M5HYy3gGeapGx5eIOxsOfxUqNE6EJq25vU3spB3daBgNdGZHlcBYl+HkcokhPSWR3bgUPwJpIMXeO62A4OXHF2PMY5UlHtBN2xhh4pmdAcvHQA3BV2muDUIWkzDVoKGfh0RjuAoJsUdSkvfSh9GAUAQFTY+8p85nZVzNRNiP0MTxaCDt4IOg1yibV0miFv121fJQA1E/Uq6mmYhOzLibqp/gZbA0EHKaIulXUusi8dAFCZsNOWwABCDiDo9RF0GtAXQjN2JMgB4Je491nYQwcvb6ZuimFByCHoIIeo02z9SGgm3XWtSAkAYKOw08Q+PvDJZgLd64qFWCOHoIPioi51sEgt6xQD0EBxr6qQD3niEUQcQNDlBF3yMAKIOgD1EXjJQlkz9eYY1etiPlrEZ2hlAEGXF3XJk5sg6gDU35MPEt8Kl94SJYUcwg2ABVHXrxcLGS7Y8wcAAABADUQdxzoCAAAANRD1F5xJDwAAAABLon6xkOOc970DAAAAoGJRb7MQSzLC2joAAABgR9hHwqJOYfgjeOwAAABA9aIe0gltC3koAtCH1w4AAM0F+9AteevK3IlNtH/9ibrZwzrNs4+ds+nJ46d/I/27EZ4WaMB4pD7fRsllAEBhI6Jf40U1XLAXv+qVxqnU5GX5g21NohKXcLlpK2DKdY8sXm9ewoqvNRcOjL2Qx97liqWsUz4aWTnQV1eNZ7rGPReicmk5PhaeZ+F8pIqusZ/nb6/IvXqBJVY/hP1IcIub1Np8x5RxcsRIri3aA0Gvn6CzkTx1vQZEzmf/wubkA4Ke6zrTdjwFOfrAxmttQVLtsrW1NdOvof5yW78G6uakpKqZ6ddEv3b1tbxH19OA8CMZ6jF6YHMmzurmmONezj5y7vit0aR0XPUkDhRimPK9/RXvfZBio483/YE7aGNnhH3Oojphz5EMz31l7rSmSN0c8HDW4LXDHs169f2PHLy2mbpdy5sI1O3a3yrlPYTt+v9x33JFzGn8nKa0XdzOU26zQL1dT/3Ekds44H/fZduwfC90fPNOQ8fxPGUctNXbB2VNU8bGrEIbH+m+eLY0qaRQ/AE5don+2k95vkOc61Ef76LDazAjXlfJE56/5N854jW3TsXX7mrIfW2I2nbI3eW2TLmuhWttleHZv1jx3NuJ95+7+pzTcl/w/G9dU2h7KWrFdQUp9nu89J7l3I7M/RAeuh/e+5Rnl28J/RrvfYoZXWZoXXUnOUsGteJhyve6aZEpHjMjfrnMyXI0gcQCfdh5Wz7Tz4meXTLUTpHCaw98hXc+gKA3R+hBedos6jiutmawV7Y86T2u49iBmHsDrYU/SAg39c89nkTup/TVzM8Vgg6aChn05PIDfX2UZzYMVnJvRdh1YkF0wpTvPa5h1AFi7s/EizxxSpBLbg1+qL83X/LO6f8HeT4bgg6aymMeMEmDT3kKz/SAO0bzlBbRNCGNXBCeNO983fKVCwWWEuu/9O999XbC1wG6nVeifqafaZQYJ+0U7zx3IhwEHTQVGkC76mYrU3JWTMmDWMpoHkcrJiHX9teB61uXGEWRjwkeoXcM2f4kbVLMtMgzxT500ORZ8pxFfRkKhb2LFqonNTrzYM5eHJaJ/LQ/5DgcrxH73EDQAQbV2+vmZPD7aJ3C0L7aNCJLordMmPF9LkFtN0vpp++ju/k9VlK+Nyk6ViDoAKJ+E9qapBhL4D9plRfTtrHRpK5b1DOqoI/Ste2kiPoeqsR5bXvSJpLPi34eBB2Am4FFBh1r5/V7rjP1dhWxcLn+ORlW9oqmDt/LXKXvwjjCkwYQdABuQ14Q9qHXj7Sw5pgrsi1HYpyOzPCkY3mCQpUk9/CYAQQdgNseUBctUZp9l05bYxFME3XaJvQicYzwpbq9N9hVBivaPEDXg6ADAN4Yfwq5DtEStXuuI7U6ozjkV5ogzh28l1nKvVBkAaF3CDoAYMlgkrGcoCVq91xpokYRmFnGX6GEOldPMDtImWz0kCDXbFBYBpgmcvQ6og3Gf0BV49SbNdUIbblWXHwRdWrDba4MR8dY3lt6C0Vo6LlHFmujRxnug8qHdtXb57sHDjx/F/rpLOW6Zp6Mn0gBAAAAoLn8vwADAIDWnwBtGEB/AAAAAElFTkSuQmCC"

def add_gibson_logo_overlay(kml):
    """Add the Gibson logo as a fixed ScreenOverlay (bottom-left corner)."""
    try:
        so = kml.newscreenoverlay(name="Gibson Logo")
        so.icon.href = "logo.png"
        so.overlayxy = OverlayXY(x=0, y=0, xunits=Units.fraction, yunits=Units.fraction)
        so.screenxy = ScreenXY(x=25, y=95, xunits=Units.pixels, yunits=Units.pixels)
        so.rotationxy = RotationXY(x=0.5, y=0.5, xunits=Units.fraction, yunits=Units.fraction)
        so.size = Size(x=300, y=0, xunits=Units.pixels, yunits=Units.pixels)
    except Exception:
        pass

# -------------------------
# KML namespace
# -------------------------
KML_NS = "http://www.opengis.net/kml/2.2"
ET.register_namespace("", KML_NS)
Q = lambda tag: "{%s}%s" % (KML_NS, tag)

# -------------------------
# Color map (KML uses aabbggrr)
# -------------------------
KML_COLOR_MAP = {
    "red": "ff0000ff",
    "blue": "ffff0000",
    "yellow": "ff00ffff",
    "purple": "ff800080",
    "green": "ff00ff00",
    "orange": "ff008cff",
    "white": "ffffffff",
    "black": "ff000000"
}

MAP_NOTE_ICON = "http://www.earthpoint.us/Dots/GoogleEarth/pal3/icon62.png"
MAP_NOTE_FALLBACK = "https://maps.google.com/mapfiles/kml/pal3/icon54.png"
RED_X_ICON = "http://maps.google.com/mapfiles/kml/pal3/icon56.png"

# -------------------------
# AGM icon whitelist (YOUR EXACT REQUIRED OPTIONS)
# -------------------------
AGM_ALLOWED_ICON_URLS = {
    "http://maps.google.com/mapfiles/kml/paddle/ylw-circle.png",   # Yellow Preliminary Location
    "http://maps.google.com/mapfiles/kml/shapes/triangle.png",     # Purple Valve
    "http://maps.google.com/mapfiles/kml/paddle/blu-circle.png",   # Blue Survey Point
    "http://maps.google.com/mapfiles/kml/shapes/flag.png",         # Red AGM
}

# -------------------------
# Helpers
# -------------------------
def safe_str(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s != "" else None

def normalize_agm_name(raw_name):
    s = safe_str(raw_name)
    if s is None:
        return ""
    if re.fullmatch(r"0+\d+", s):
        return s
    if re.fullmatch(r"\d+", s):
        if len(s) >= 4:
            return s
        if len(s) < 3:
            return s.zfill(3)
        return s
    try:
        f = float(s)
        if f.is_integer():
            i = int(f)
            s_digits = str(i)
            if len(s_digits) < 3:
                return s_digits.zfill(3)
            return s_digits
    except:
        pass
    return s

def normalize_color_value(val):
    c = safe_str(val)
    if not c:
        return None
    cl = c.lower()
    if cl in KML_COLOR_MAP:
        return KML_COLOR_MAP[cl]
    if len(c) == 8 and all(ch in "0123456789abcdefABCDEF" for ch in c):
        return c.lower()
    return None

def set_icon(point, href_value):
    href = safe_str(href_value)
    if not href:
        return
    try:
        point.style.iconstyle.icon.href = href
    except:
        pass

def set_icon_color(point, color_value):
    col = normalize_color_value(color_value)
    if not col:
        return
    try:
        point.style.iconstyle.color = col
    except:
        pass

def set_linestring_style(ls, color_value):
    col = normalize_color_value(color_value)
    if not col:
        return
    try:
        ls.style.linestyle.color = col
        ls.style.linestyle.width = 3
    except:
        pass

def choose_note_icon_href(icon_value):
    v = safe_str(icon_value)
    if v is None:
        return None
    vl = v.lower()
    if vl == "map note":
        return MAP_NOTE_ICON
    if vl == "red x":
        return RED_X_ICON
    return v

def haversine_m(lat1, lon1, lat2, lon2):
    R = 6371000.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dl/2)**2
    return 2 * R * math.asin(math.sqrt(a))

# -------------------------
# Line builder (prevents "looping back" by splitting on big jumps)
# -------------------------
def add_lines_with_autosplit(folder, df, color_col="LineStringColor", split_jump_m=5000.0):
    if df is None or df.empty:
        return False

    chosen_color = None
    if color_col in df.columns:
        non_null = df[color_col].dropna().astype(str).str.strip()
        if len(non_null) > 0:
            chosen_color = non_null.iloc[0]

    created_any = False
    seg = []
    prev = None

    def flush(segment):
        nonlocal created_any
        if len(segment) < 2:
            return
        ls = folder.newlinestring()
        ls.coords = segment
        if chosen_color:
            set_linestring_style(ls, chosen_color)
        created_any = True

    for _, row in df.iterrows():
        lat = row.get("Latitude")
        lon = row.get("Longitude")
        if pd.isna(lat) or pd.isna(lon):
            flush(seg)
            seg = []
            prev = None
            continue

        try:
            pt = (float(lon), float(lat))
        except:
            continue

        if prev is not None:
            if pt == prev:
                continue
            jump = haversine_m(prev[1], prev[0], pt[1], pt[0])
            if jump > split_jump_m:
                flush(seg)
                seg = []

        seg.append(pt)
        prev = pt

    flush(seg)
    return created_any

# -------------------------
# Placemark creators
# -------------------------
def add_agm_point(folder, row):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return False
    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except:
        return False

    p = folder.newpoint()
    name_val = normalize_agm_name(row.get("Name"))
    p.name = str(name_val)
    p.description = str(name_val)
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"
    except:
        pass
    p.coords = [(lon_f, lat_f)]

    # AGM icon: MUST be one of your exact URLs
    icon_raw = safe_str(row.get("Icon"))
    if icon_raw:
        icon_norm = icon_raw.strip()
        if icon_norm in AGM_ALLOWED_ICON_URLS:
            set_icon(p, icon_norm)
        else:
            # If it's not one of the four, do nothing (prevents unexpected Earthpoint substitutions)
            pass

    # Tint by IconColor (Yellow/Purple/Blue/Red)
    set_icon_color(p, row.get("IconColor"))
    return True

def add_access_point(folder, row):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return False
    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except:
        return False

    p = folder.newpoint()
    name_val = safe_str(row.get("Name")) or ""
    p.name = str(name_val)
    p.description = str(name_val)
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"
    except:
        pass
    p.coords = [(lon_f, lat_f)]

    if "icon" in row.index:
        set_icon(p, row.get("icon"))
    elif "Icon" in row.index:
        set_icon(p, row.get("Icon"))
    return True

def add_note_point(folder, row):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return ""
    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except:
        return ""

    p = folder.newpoint()
    name_val = safe_str(row.get("Name")) or ""
    name_str = str(name_val)
    p.name = name_str
    p.description = name_str
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"
    except:
        pass
    p.coords = [(lon_f, lat_f)]

    href = choose_note_icon_href(row.get("Icon"))
    if href:
        try:
            p.style.iconstyle.icon.href = href
        except:
            try:
                p.style.iconstyle.icon.href = MAP_NOTE_FALLBACK
            except:
                pass
    return href or ""

# -------------------------
# KML post-process: StyleMaps for Notes ONLY (hide until hover when flagged)
# -------------------------
def inject_hover_stylemaps_for_notes_with_flags(kml_bytes, notes_flags_by_name, notes_folder_name="Notes"):
    root = ET.fromstring(kml_bytes)
    doc = root.find(".//" + Q("Document"))
    if doc is None:
        if root.tag == Q("Document"):
            doc = root
        else:
            return kml_bytes

    notes_folder = None
    for folder in doc.findall(Q("Folder")):
        nm = folder.find(Q("name"))
        if nm is not None and nm.text and nm.text.strip().lower() == notes_folder_name.lower():
            notes_folder = folder
            break
    if notes_folder is None:
        for folder in doc.findall(Q("Folder")):
            nm = folder.find(Q("name"))
            if nm is not None and nm.text and nm.text.strip().lower() == "notes":
                notes_folder = folder
                break
    if notes_folder is None:
        return kml_bytes

    style_by_id = {}
    for st in doc.findall(Q("Style")):
        sid = st.get("id")
        if sid:
            style_by_id[sid] = st

    def href_from_style(style_el):
        if style_el is None:
            return None
        href_el = style_el.find(".//" + Q("Icon") + "/" + Q("href"))
        if href_el is not None and href_el.text:
            return href_el.text.strip()
        return None

    def href_from_pm(pm):
        href_el = pm.find(".//" + Q("Icon") + "/" + Q("href"))
        if href_el is not None and href_el.text:
            return href_el.text.strip()
        su = pm.find(Q("styleUrl"))
        if su is not None and su.text and su.text.strip().startswith("#"):
            sid = su.text.strip()[1:]
            return href_from_style(style_by_id.get(sid))
        return None

    def get_name(pm):
        n = pm.find(Q("name"))
        return (n.text or "").strip() if n is not None else ""

    pairs = []
    pm_info = []
    for pm in notes_folder.findall(Q("Placemark")):
        name = get_name(pm)
        hide_flag = bool(notes_flags_by_name.get(name, True))
        href = href_from_pm(pm) or MAP_NOTE_FALLBACK
        pm_info.append((pm, href, hide_flag))
        key = (href, hide_flag)
        if key not in pairs:
            pairs.append(key)

    if not pairs:
        return kml_bytes

    first_folder = doc.find(Q("Folder"))

    def insert_before_first_folder(el):
        if first_folder is None:
            doc.append(el)
        else:
            idx = list(doc).index(first_folder)
            doc.insert(idx, el)

    key_to_smid = {}
    for i, (href, hide_flag) in enumerate(pairs, start=1):
        sm_id = f"sm_notes_{i}"

        st_n = ET.Element(Q("Style"), {"id": f"{sm_id}_normal"})
        is_n = ET.SubElement(st_n, Q("IconStyle"))
        ic_n = ET.SubElement(is_n, Q("Icon"))
        ET.SubElement(ic_n, Q("href")).text = href
        ls_n = ET.SubElement(st_n, Q("LabelStyle"))
        if hide_flag:
            ET.SubElement(ls_n, Q("scale")).text = "0.01"
            ET.SubElement(ls_n, Q("color")).text = "00ffffff"
        else:
            ET.SubElement(ls_n, Q("scale")).text = "1"
            ET.SubElement(ls_n, Q("color")).text = "ffffffff"

        st_h = ET.Element(Q("Style"), {"id": f"{sm_id}_highlight"})
        is_h = ET.SubElement(st_h, Q("IconStyle"))
        ic_h = ET.SubElement(is_h, Q("Icon"))
        ET.SubElement(ic_h, Q("href")).text = href
        ls_h = ET.SubElement(st_h, Q("LabelStyle"))
        ET.SubElement(ls_h, Q("scale")).text = "1"
        ET.SubElement(ls_h, Q("color")).text = "ffffffff"

        sm = ET.Element(Q("StyleMap"), {"id": sm_id})
        p1 = ET.SubElement(sm, Q("Pair"))
        ET.SubElement(p1, Q("key")).text = "normal"
        ET.SubElement(p1, Q("styleUrl")).text = f"#{sm_id}_normal"
        p2 = ET.SubElement(sm, Q("Pair"))
        ET.SubElement(p2, Q("key")).text = "highlight"
        ET.SubElement(p2, Q("styleUrl")).text = f"#{sm_id}_highlight"

        insert_before_first_folder(st_n)
        insert_before_first_folder(st_h)
        insert_before_first_folder(sm)
        key_to_smid[(href, hide_flag)] = sm_id

    for pm, href, hide_flag in pm_info:
        smid = key_to_smid.get((href, hide_flag))
        if not smid:
            continue
        for existing in pm.findall(Q("styleUrl")):
            pm.remove(existing)
        for inline_style in pm.findall(Q("Style")):
            pm.remove(inline_style)
        ET.SubElement(pm, Q("styleUrl")).text = f"#{smid}"

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


# ===========================================================================
# PART 2 : KMZ -> DMT  (library code from the KMZ-To-DMT Generator)
# ===========================================================================
# KMZ To DMT Generator  --  single-file Streamlit app.
# Upload a KMZ from the XLSX-To-KMZ Generator; download a DeLorme Street Atlas
# .dmt with four layers: AGMs, Access, Centerline, Notes.





FREESECT   = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT    = 0xFFFFFFFD
DIFSECT    = 0xFFFFFFFC
NOSTREAM   = 0xFFFFFFFF
SIG = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'


class DirEntry:
    __slots__ = ("raw", "name", "type", "color", "left", "right", "child",
                 "start", "size")

    def __init__(self, raw):
        self.raw = bytearray(raw)
        nlen = struct.unpack_from('<H', raw, 64)[0]
        self.name = raw[:max(0, nlen - 2)].decode('utf-16-le') if nlen else ''
        self.type = raw[66]
        self.color = raw[67]
        self.left = struct.unpack_from('<I', raw, 68)[0]
        self.right = struct.unpack_from('<I', raw, 72)[0]
        self.child = struct.unpack_from('<I', raw, 76)[0]
        self.start = struct.unpack_from('<I', raw, 116)[0]
        self.size = struct.unpack_from('<I', raw, 120)[0]

    def set_name(self, name):
        enc = name.encode('utf-16-le')
        assert len(enc) <= 62, "name too long"
        self.raw[0:64] = b'\x00' * 64
        self.raw[0:len(enc)] = enc
        struct.pack_into('<H', self.raw, 64, len(enc) + 2)
        self.name = name

    def pack(self):
        r = bytearray(self.raw)
        r[67] = self.color & 0xFF
        struct.pack_into('<I', r, 68, self.left & 0xFFFFFFFF)
        struct.pack_into('<I', r, 72, self.right & 0xFFFFFFFF)
        struct.pack_into('<I', r, 76, self.child & 0xFFFFFFFF)
        struct.pack_into('<I', r, 116, self.start & 0xFFFFFFFF)
        struct.pack_into('<I', r, 120, self.size & 0xFFFFFFFF)
        struct.pack_into('<I', r, 124, 0)
        return bytes(r)


class OLE:
    def __init__(self, data):
        data = bytes(data)
        assert data[:8] == SIG, "not an OLE compound file"
        self.data = data
        self.sect_size = 1 << struct.unpack_from('<H', data, 30)[0]
        self.mini_size = 1 << struct.unpack_from('<H', data, 32)[0]
        self.dir_start = struct.unpack_from('<I', data, 48)[0]
        self.mini_cutoff = struct.unpack_from('<I', data, 56)[0]
        self.minifat_start = struct.unpack_from('<I', data, 60)[0]
        self.difat_start = struct.unpack_from('<I', data, 68)[0]
        self.num_difat = struct.unpack_from('<I', data, 72)[0]
        self._read_difat()
        self._read_fat()
        self._read_dir()
        self._read_minifat()
        root = self.entries[0]
        self.ministream = self._read_chain(root.start, root.size)

    def _sector(self, i):
        off = 512 + i * self.sect_size
        return self.data[off:off + self.sect_size]

    def _read_difat(self):
        self.difat = []
        for i in range(109):
            v = struct.unpack_from('<I', self.data, 76 + i * 4)[0]
            if v == FREESECT:
                break
            self.difat.append(v)
        sec, n = self.difat_start, self.num_difat
        while sec not in (ENDOFCHAIN, FREESECT) and n > 0:
            s = self._sector(sec)
            for i in range((self.sect_size // 4) - 1):
                v = struct.unpack_from('<I', s, i * 4)[0]
                if v != FREESECT:
                    self.difat.append(v)
            sec = struct.unpack_from('<I', s, self.sect_size - 4)[0]
            n -= 1

    def _read_fat(self):
        self.fat = []
        for sec in self.difat:
            s = self._sector(sec)
            for i in range(self.sect_size // 4):
                self.fat.append(struct.unpack_from('<I', s, i * 4)[0])

    def _chain_sectors(self, start):
        out, s = [], start
        while s not in (ENDOFCHAIN, FREESECT):
            out.append(s)
            s = self.fat[s]
            if len(out) > 2_000_000:
                break
        return out

    def _read_chain(self, start, size):
        return b''.join(self._sector(s) for s in self._chain_sectors(start))[:size]

    def _read_dir(self):
        raw = b''.join(self._sector(s) for s in self._chain_sectors(self.dir_start))
        self.entries = []
        for i in range(0, len(raw), 128):
            e = raw[i:i + 128]
            if len(e) < 128:
                break
            self.entries.append(DirEntry(e))

    def _read_minifat(self):
        if self.minifat_start in (ENDOFCHAIN, FREESECT):
            self.minifat = []
            return
        raw = b''.join(self._sector(s) for s in self._chain_sectors(self.minifat_start))
        self.minifat = [struct.unpack_from('<I', raw, i)[0] for i in range(0, len(raw), 4)]

    def _mini_chain(self, start):
        out, s = [], start
        while s not in (ENDOFCHAIN, FREESECT):
            out.append(s)
            s = self.minifat[s]
            if len(out) > 2_000_000:
                break
        return out

    def read(self, entry):
        if entry.type == 2 and entry.size < self.mini_cutoff:
            return b''.join(
                self.ministream[s * self.mini_size:(s + 1) * self.mini_size]
                for s in self._mini_chain(entry.start))[:entry.size]
        return self._read_chain(entry.start, entry.size)

    def stream_map(self):
        return {e.name: self.read(e) for e in self.entries if e.type == 2}



def _rebuild_directory(entries):
    def child_list(idx):
        out = []
        def inorder(i):
            if i == NOSTREAM or i >= len(entries):
                return
            inorder(entries[i].left)
            out.append(i)
            inorder(entries[i].right)
        inorder(entries[idx].child)
        return out

    # discover parent -> children from the existing tree
    groups = {}
    stack = [0]
    while stack:
        p = stack.pop()
        kids = child_list(p)
        if kids:
            groups[p] = kids
            stack.extend(kids)

    def key(i):
        nm = entries[i].name.upper()
        return (len(entries[i].name), nm)

    def build(idxs):
        if not idxs:
            return NOSTREAM
        mid = len(idxs) // 2
        root = idxs[mid]
        entries[root].left = build(idxs[:mid])
        entries[root].right = build(idxs[mid + 1:])
        entries[root].color = 1  # black
        return root

    for parent, kids in groups.items():
        kids = sorted(kids, key=key)
        entries[parent].child = build(kids)
        entries[parent].color = 1


def write_ole(template, new_streams, rename=None, sect_size=512, mini_size=64, mini_cutoff=4096):
    entries = [DirEntry(bytes(e.raw)) for e in template.entries]

    contents = {}
    for idx, e in enumerate(entries):
        contents[idx] = new_streams.get(e.name, template.read(e)) if e.type == 2 else b''

    # apply renames AFTER content resolution (which keys off original names)
    if rename:
        for e in entries:
            if e.name in rename:
                e.set_name(rename[e.name])
    _rebuild_directory(entries)

    # mini stream: streams smaller than cutoff
    minifat, ministream = [], bytearray()
    for idx, e in enumerate(entries):
        if e.type != 2:
            continue
        data = contents[idx]
        if 0 < len(data) < mini_cutoff:
            nsect = (len(data) + mini_size - 1) // mini_size
            first = len(minifat)
            for j in range(nsect):
                minifat.append(first + j + 1 if j < nsect - 1 else ENDOFCHAIN)
            ministream += data + b'\x00' * (nsect * mini_size - len(data))
            e.start, e.size = first, len(data)
        elif len(data) == 0:
            e.start, e.size = ENDOFCHAIN, 0

    root = entries[0]
    big = [(idx, contents[idx]) for idx, e in enumerate(entries)
           if e.type == 2 and len(contents[idx]) >= mini_cutoff]

    dir_sects = (len(entries) * 128 + sect_size - 1) // sect_size
    minifat_bytes = b''.join(struct.pack('<I', v) for v in minifat)
    minifat_sects = (len(minifat_bytes) + sect_size - 1) // sect_size if minifat else 0
    ministream_sects = (len(ministream) + sect_size - 1) // sect_size if ministream else 0
    big_counts = [(len(d) + sect_size - 1) // sect_size for _, d in big]
    non_fat = dir_sects + minifat_sects + ministream_sects + sum(big_counts)

    fat_sects = 1
    while True:
        needed = ((non_fat + fat_sects) * 4 + sect_size - 1) // sect_size
        if needed <= fat_sects:
            break
        fat_sects = needed

    fat = [FREESECT] * (fat_sects * (sect_size // 4))
    nxt = 0

    def take(n):
        nonlocal nxt
        ids = list(range(nxt, nxt + n))
        nxt += n
        return ids

    def link(ids):
        for k, s in enumerate(ids):
            fat[s] = ids[k + 1] if k < len(ids) - 1 else ENDOFCHAIN

    fat_ids = take(fat_sects)
    dir_ids = take(dir_sects)
    minifat_ids = take(minifat_sects)
    ministream_ids = take(ministream_sects)
    for s in fat_ids:
        fat[s] = FATSECT
    link(dir_ids); link(minifat_ids); link(ministream_ids)

    big_ids = []
    for (idx, data), cnt in zip(big, big_counts):
        ids = take(cnt)
        link(ids)
        entries[idx].start, entries[idx].size = (ids[0] if ids else ENDOFCHAIN), len(data)
        big_ids.append(ids)

    root.start = ministream_ids[0] if ministream_ids else ENDOFCHAIN
    root.size = len(ministream)

    sectors = [b'\x00' * sect_size] * nxt

    def put(ids, data):
        padded = data + b'\x00' * (len(ids) * sect_size - len(data))
        for k, s in enumerate(ids):
            sectors[s] = padded[k * sect_size:(k + 1) * sect_size]

    dir_bytes = b''.join(e.pack() for e in entries)
    if len(dir_bytes) % sect_size:
        dir_bytes += b'\x00' * (sect_size - len(dir_bytes) % sect_size)
    put(dir_ids, dir_bytes)
    if minifat_sects:
        put(minifat_ids, minifat_bytes)
    if ministream_sects:
        put(ministream_ids, bytes(ministream))
    for (idx, data), ids in zip(big, big_ids):
        put(ids, data)
    put(fat_ids, b''.join(struct.pack('<I', v) for v in fat))

    hdr = bytearray(512)
    hdr[0:8] = SIG
    struct.pack_into('<H', hdr, 24, 0x003E)
    struct.pack_into('<H', hdr, 26, 0x0003)
    struct.pack_into('<H', hdr, 28, 0xFFFE)
    struct.pack_into('<H', hdr, 30, 9)
    struct.pack_into('<H', hdr, 32, 6)
    struct.pack_into('<I', hdr, 44, fat_sects)
    struct.pack_into('<I', hdr, 48, dir_ids[0])
    struct.pack_into('<I', hdr, 56, mini_cutoff)
    struct.pack_into('<I', hdr, 60, minifat_ids[0] if minifat_sects else ENDOFCHAIN)
    struct.pack_into('<I', hdr, 64, minifat_sects)
    struct.pack_into('<I', hdr, 68, ENDOFCHAIN)
    for i in range(109):
        struct.pack_into('<I', hdr, 76 + i * 4, fat_ids[i] if i < len(fat_ids) else FREESECT)

    return bytes(hdr) + b''.join(sectors)




KNS = {'k': 'http://www.opengis.net/kml/2.2'}


def _ln(el):
    return el.tag.split('}')[-1]


def _text(el, path):
    f = el.find(path, KNS)
    return f.text.strip() if f is not None and f.text else None


def _coords(s):
    out = []
    for tok in s.replace('\n', ' ').split():
        parts = tok.split(',')
        if len(parts) >= 2:
            out.append((float(parts[0]), float(parts[1])))  # (lon, lat)
    return out


class Placemark:
    def __init__(self, name, kind, coords, icon=None, color=None):
        self.name = name or ''
        self.kind = kind          # 'point' or 'line'
        self.coords = coords      # list of (lon, lat)
        self.icon = icon          # e.g. 'flag.png'
        self.color = color        # KML aabbggrr hex string


def _read_kml_bytes(path_or_bytes):
    if isinstance(path_or_bytes, (bytes, bytearray)):
        zf = zipfile.ZipFile(__import__('io').BytesIO(path_or_bytes))
    else:
        zf = zipfile.ZipFile(path_or_bytes)
    name = next((n for n in zf.namelist() if n.lower().endswith('.kml')), None)
    if name is None:
        raise ValueError("No .kml inside KMZ")
    return zf.read(name)


def parse_kmz(path_or_bytes):
    raw = _read_kml_bytes(path_or_bytes)
    root = ET.fromstring(raw)

    # style id -> (icon file, color)
    styles = {}
    for s in root.iter():
        if _ln(s) == 'Style':
            sid = s.get('id')
            href = s.find('.//k:Icon/k:href', KNS)
            color = s.find('.//k:color', KNS)
            styles[sid] = (
                href.text.split('/')[-1] if href is not None and href.text else None,
                color.text if color is not None else None,
            )

    def style_of(pm):
        su = pm.find('k:styleUrl', KNS)
        if su is not None and su.text:
            return styles.get(su.text.lstrip('#'), (None, None))
        href = pm.find('.//k:Icon/k:href', KNS)
        color = pm.find('.//k:color', KNS)
        return (href.text.split('/')[-1] if href is not None and href.text else None,
                color.text if color is not None else None)

    folders = {}

    # Any Folder anywhere in the tree that directly holds placemarks is a layer.
    for f in root.iter():
        if _ln(f) != 'Folder':
            continue
        fname = _text(f, 'k:name') or 'Folder'
        items = []
        for pm in f.findall('k:Placemark', KNS):
            nm = _text(pm, 'k:name')
            icon, color = style_of(pm)
            pt = pm.find('.//k:Point/k:coordinates', KNS)
            ls = pm.find('.//k:LineString/k:coordinates', KNS)
            if pt is not None and pt.text:
                items.append(Placemark(nm, 'point', _coords(pt.text), icon, color))
            elif ls is not None and ls.text:
                items.append(Placemark(nm, 'line', _coords(ls.text), icon, color))
        if items:
            folders[fname] = items
    return folders



_TEMPLATE_B64 = """
0M8R4KGxGuEAAAAAAAAAAAAAAAAAAAAAPgADAP7/CQAGAAAAAAAAAAAAAAABAAAAAQAAAAAAAAAAEAAACAAAAAEAAAD+////AAAA
AAAAAAD/////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////9////AgAAAAMAAAAE
AAAABQAAAAYAAAAHAAAA/v////7///8KAAAACwAAAAwAAAANAAAA/v///w8AAAAQAAAAEQAAABIAAAATAAAAFAAAABUAAAAWAAAA
FwAAABgAAAAZAAAAGgAAABsAAAAcAAAAHQAAAB4AAAAfAAAAIAAAACEAAAAiAAAAIwAAACQAAAAlAAAAJgAAACcAAAAoAAAAKQAA
ACoAAAArAAAALAAAAC0AAAAuAAAALwAAADAAAAAxAAAAMgAAADMAAAA0AAAANQAAADYAAAA3AAAAOAAAADkAAAA6AAAAOwAAADwA
AAA9AAAAPgAAAD8AAABAAAAAQQAAAEIAAABDAAAARAAAAEUAAABGAAAA/v//////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////1IAbwBvAHQAIABFAG4AdAByAHkAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWAAUB//////////8BAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAOCvyeI20dwBCQAAAMAIAAAAAAAARABlAEwAbwByAG0AZQBDAG8AbQBwAG8AbgBlAG4AdABzAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAACQAAQH//////////xQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOA+x+I20dwB4K/J4jbR3AEAAAAAAAAA
AAAAAABEAGUATABvAHIAbQBlAC4AQQBuAG4AbwB0AGEAdABlAC4AVwBvAHIAawBzAHAAYQBjAGUAAAAAAAAAAAAAAAAANgABAf//
////////AwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4D7H4jbR3AHgr8niNtHcAQAAAAAAAAAAAAAAAEEAbgBuAG8AdABhAHQAZQAu
AEYAaQBsAGUAbgBhAG0AZQBzAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAmAAIBBwAAAAYAAAD/////AAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJIAAAAAAAAAQwBlAG4AdABlAHIAbABpAG4AZQA0ADEAIAAoADQAKQAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACIAAgH///////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAADAAAAkwAAAAAAAABuAG8AdABlAHMAIAAoADIANgApAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAFgACAf///////////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYAAACzAAAAAAAAAEMAbwBtAGIA
aQBuAGUAZAAgAEEAYwBjAGUAcwBzACAARgBpAGwAZQBzADEAIAAoADEAMQApAAAAAAAAAAAAAAA4AAIBCAAAAP//////////AAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACQAAAJMAAAAAAAAARgBpAG4AYQBsACAAQQBHAE0AcwA2ADgAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoAAgEFAAAABAAAAP////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAOAAAAj3EAAAAAAABBAG4AbgBvAHQAYQB0AGUALgBBAGMAdABpAHYAZQBGAGkAbABlAG4AYQBtAGUAcwAAAAAA
AAAAAAAAAAAAAAAAMgACAf///////////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwAAAAtAAAAAAAA
AEQAZQBMAG8AcgBtAGUALgBBAHUAdABvAE4AYQB2AC4AVwBvAHIAawBzAHAAYQBjAGUAAAAAAAAAAAAAAAAAAAA0AAEBFwAAAAIA
AAAKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADgr8niNtHcAeCvyeI20dwBAAAAAAAAAAAAAAAARABlAEwAbwByAG0AZQBBAHUAdABv
AE4AYQB2AEYAaQBsAGUAcwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACgAAgH///////////////8AAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAANAAAAGAAAAAAAAABEAGUATABvAHIAbQBlAC4ATQBhAHAAMgBEAC4AMQAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAIAABAf//////////DQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4K/J4jbR3AHgr8niNtHcAQAA
AAAAAAAAAAAAAE0AYQBwADIARABTAHQAYQB0AGUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAW
AAIB////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADgAAADYAAAAAAAAATQBhAHAAMgBEAFMA
dABhAHQAZQAyAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABgAAgEMAAAA//////////8AAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPAAAANgAAAAAAAABEAGUATABvAHIAbQBlAC4AUAByAGUAZgBlAHIAZQBuAGMA
ZQBzAC4AMQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALAABAf//////////DwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4K/J4jbR3AHg
r8niNtHcAQAAAAAAAAAAAAAAAFAAcgBlAGYAQwB1AHIAcgBlAG4AdABTAGUAdAB0AGkAbgBnAHMAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAoAAIB////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAKsDAAAAAAAARABl
AEwAbwByAG0AZQAuAFAAcgBpAG4AdAAyAC4AMQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACIAAQELAAAADgAAABEA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAOCvyeI20dwB4K/J4jbR3AEAAAAAAAAAAAAAAABQAGEAZwBlAEkAbgBmAG8AAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEgACARMAAAASAAAA/////wAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAB8AAAAoAAAAAAAAAE8AYgBqAGUAYwB0AEMAbwB1AG4AdAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAYAAIB////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAgA
AAAAAAAATwBiAGoAZQBjAHQAcwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAQH/
//////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAOCvyeI20dwB4K/J4jbR3AEAAAAAAAAAAAAAAABEAGUATABvAHIAbQBlAC4A
VABhAHIAZwBlAHQAVABvAG8AbABzAC4AMQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALAABARAAAAAJAAAAFQAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAA4K/J4jbR3AHgr8niNtHcAQAAAAAAAAAAAAAAAFQAbwBvAGwAQwBsAGEAcwBzAFMAdABvAHIAYQBnAGUAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAiAAEB//////////8WAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADgr8niNtHcAeCvyeI2
0dwBAAAAAAAAAAAAAAAATQBhAHAAMgBEAE0AZQBhAHMAdQByAGUAVABvAG8AbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAACIAAgH///////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhAAAADAAAAAAAAABEAGUATABv
AHIAbQBlAC4AWABEAGEAdABhAC4AVwBvAHIAawBzAHAAYQBjAGUAAAAAAAAAAAAAAAAAAAAAAAAAMAABAf//////////GAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAA4K/J4jbR3AHgr8niNtHcAQAAAAAAAAAAAAAAAFgARABhAHQAYQAuAEYAaQBsAGUAbgBhAG0AZQBz
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAIB////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAIgAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAD///////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////
/////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////////////////AAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAIAAAD+////BAAAAAUAAAD+////BwAAAAgAAAD+////CgAAAAsA
AAD+/////v////7////+/////v///xEAAAASAAAAEwAAABQAAAAVAAAAFgAAABcAAAAYAAAAGQAAABoAAAAbAAAAHAAAAB0AAAAe
AAAA/v////7////+/////v////7///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAJQAAAEM6XERlTG9ybWUgRG9jc1xEcmF3XEZpbmFsIEFHTXM2OC5hbjEEAAAAEAAAAENl
bnRlcmxpbmU0MSAoNCkBAAAACgAAAG5vdGVzICgyNikBAAAAGwAAAENvbWJpbmVkIEFjY2VzcyBGaWxlczEgKDExKQEAAAAMAAAA
RmluYWwgQUdNczY4AQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACPAAAAJS0AAAAAAAAC
AAAAAAACAAEAAAAOAABBEScAAAAAAAACAAIAAAAXAAABAAAAAADq6olpAAAAAOrqiWkAAAAAAgAAAAAA/wAAAAMAAAAAAAAAAAAA
AAAAAAACAAIAAAABAAAAAADUGcdQU6ZwbAAAAQAAAAAAvBrHUGmlcGwAAAQAAAAAAAAAAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACvAAAAJS0AAAAAAAACAAEAAAABAAAAAAB77VVRVPLlawYAzv///xQAAAANAABB
1icAAA4BBAAAAAAAAAAAACIAQnVja2V5ZSBLZXkAAQAAAAAAgGLhaQAAAACFYuFpAAAAAAUAQXJpYWwAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAKAAAAAAAAAAAAAAADAAAA////AAAAAAADAAAAAgAAAAAABAAAAAAAAAADAAAAAAAAAAAAAAAAAAAAAAAAAACPAAAA
JS0AAAAAAAACAAAAAAACAAEAAAAPAABBJnYAAAAAAAACAAIAAAAXAAABAAAAAAATT+JpAAAAABNP4mkAAAAAAgAAAAAAAAD/AAMA
AAAAAAAAAAAAAAAAAAACAAIAAAABAAAAAABdqORQAFIfbAAAAQAAAAAAVKjkUBFSH2wAAAQAAAAAAAAAAwAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAJQAAAEM6XERlTG9ybWUgRG9jc1xEcmF3XEZpbmFsIEFH
TXM2OC5hbjEAAAAAAAAAAAAAAAAAAAAAAAAAAwAAAAAAAAAtMgwAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAIAAABosEJRALAJbAYAAABVVVVVVVX1PwAAAAAAAPA/AAAAAAAAAIABEQAAAGiwQlEAsAlsAAAAAAAA
AAAAAAACAAAAaLBCUQCwCWwGAAAAVVVVVVVV9T8AAAAAAADwPwAAAAAAAACAAREAAABosEJRALAJbAAAAAAAAAAAAAAABAAAAAYA
AAB0JwAAAAAAABQnAAABAAAAEycAAAEAAAASJwAAAQAAABEnAAABAAAAJCcAACUnAAAAAAAAtQAAAEJhc2UgTWFwLEJ1c2luZXNz
IFBvaW50cyBvZiBJbnRlcmVzdCAoTWFqb3IpLENvdW50eSBCb3JkZXJzLEV4aXRzLEludGVybmF0aW9uYWwgTGFiZWxzLFBsYWNl
cyAoTWlub3IpLFBvaW50cyBvZiBJbnRlcmVzdCAoTWFqb3IpLFJvYWRzIChNaW5vciksVXJiYW4gQXJlYSBDb2xvciwtRE1MX0NV
U1RPTV9GSUxURVJfSUQAAAAAAAAAAFY0ErCchuZAAAAAAE5iPkE9AAAAEzyBCBQ8gQgWPIEIFzyBCCE8gQgiPIEIJDyBCCU8gQgn
PIEIMDyBCEA8gQhQPIEIYDyBCJA8gQjyPIEI8zyBCPc8gQj4PIEI/DyBCP08gQj/PIEIAT2BCAQ9gQggPYEIUD2BCHE9gQhyPYEI
cz2BCHQ9gQh1PYEIdj2BCHc9gQjgPYEIAj6BCCBEgQgwRIEIUESBCGBEgQgASIEIAEyBCAAYAhEADAUREBAFEQAADhEACA8REBSD
EQBQgxEQBIURIASFETAEhRFABIURUASFEXAEhRGABIURAACKEVEIixFgCIsREQSMETEEjBEAAI8RAACEGEYAAABTb2Z0d2FyZVxE
ZUxvcm1lXFNBMjAxNVxEZUxvcm1lQ29tcG9uZW50c1xEZUxvcm1lLlNBMjAxNS5QcmVmZXJlbmNlcy4xIwAAAERlTG9ybWUgU3Ry
ZWV0IEF0bGFzIFVTQa4gMjAxNSBQbHVzJQEAAEdyaWRzLCBNYXAgQ2VudGVyIENyb3NzaGFpciwgRXhpdHMsIE9uZSBXYXlzLCBQ
bGFjZXMgKE1pbm9yKSwgUm9hZHMgKE1pbm9yKSwgUG9pbnRzIG9mIEludGVyZXN0IChNYWpvciksIFBvaW50cyBvZiBJbnRlcmVz
dCAoTWlub3IpLCBCdXNpbmVzcyBQb2ludHMgb2YgSW50ZXJlc3QgKE1ham9yKSwgQnVzaW5lc3MgUG9pbnRzIG9mIEludGVyZXN0
IChNaW5vciksIFRvd24gQm9yZGVycywgQ291bnR5IEJvcmRlcnMsIEludGVybmF0aW9uYWwgTGFiZWxzLCBVcmJhbiBBcmVhIENv
bG9yLCBaSVAtUG9zdGFsIENvZGVzCAAAAEJhc2UgTWFwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAdQAAAAAAAACNAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAawAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAi3EAACUtAAADAAAASUwBAQMAAwABABgAGAD/////IAD//////////0JNNgAA
AAAAAAA2AAAAKAAAAGAAAAAwAAAAAQAYAAAAAAAANgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////////////////////////////
//////////////////////////////////////////////////////////8A//8A//8A//////////////////8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A/////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP////8A//8A//8A//8A/////xQUGhoaGf////8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A/////wAAALAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAg
cLAgcLAgcLAgcAAAAP////8A//8A//8A//8A/////xkZHxkZGf////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAP8A//8A/////wAAALAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcAAAAP//
//8A//8A//8A//8A//8A/////xoaIBkZGP////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A//8A
//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AP8A//8A/////wAAALAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcAAAAP////8A//8A//8A//8A
//8A/////xoaIBkZGP////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A/wAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A////
/wAAALAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcAAAAP////8A//8A//8A//8A//8A//8A/////wAAAAAA
AP////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
/wAAAAAAAJAAEowAFIYAF34AGncAHQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A/////wAAALAgcLAgcLAg
cLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcAAAAP////8A//8A//8A//8A//8A//8A/////wAAAAAAAP////8A//8A//8A
//8A//////////////////////////////////////////////8A//8A//8A//8A//8A//8A//8A//8A//8A/6oABqsABakABqUA
CJwADJIAEYcAFnwAGnMAHgAAAAAAAAAAAP8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A/////wAAALAgcLAgcLAgcLAgcLAgcLAg
cLAgcLAgcLAgcLAgcLAgcAAAAP////8A//8A//8A//8A//8A//8A//8A/////wAAAAAAAP///////////////////////x8fHgAA
AAAAAQAABAAACgAADwAADwAAACQkI/////8A//8A//8A//8A//8A//8A//8A//8A/78AAMUAAMcAAMUAAL8AALYAAKkABZoADYwA
FH4AGnMAHv8A//8A//8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A/////wAAALAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAg
cLAgcAAAAP////8A//8A//8A//8A//8A//8A//8A/////wAAAAAAAB0dEycnGSQkEyQkDycnDx8fCgUFAAAANgAAhgAAmAAAvwAA
3wAA6gAAVCQkEf////8A//8A//8A//8A//8A//8A//8A/88AANkAAN8AAOEAAN8AANkAAM8AAMEAALAAA54AC4wAFHwAG3MAHv8A
//8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A//8A/////wAAALAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcAAAAP////8A//8A
//8A//8A//8A//8A//8A//8A/////wAAAAAAAAAAJQAAPgAATwAAZAAAaQAAawAAegAAiwAApQAAxgAA7QAA/wAA/wAAaCQkDv//
//8A//8A//8A//8A//8A//8A//8A/+QAAO4AAPQAAPUAAPUAAO8AAOUAANcAAMUAALAAApsADIcAFncAHP8A//8A//8A//8A//8A
//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAP8A//8A//8A//8A//8A/////wAAALAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcLAgcAAAAP////8A//8A//8A//8A//8A//8A
//8A//8A/////wAAAAAAAAAAfgAAvwAA2gAA+QAA/wAA/wAA/wAAwwAAiwAAsQAA5AAA/wAA/wAAaCQkDv////8A//8A//8A//8A
//8A//8A/+gAAPYAAP0AAP8FBf8JCP8GBf4AAPUAAOkAANcAAMEAAKkABZIAEX8AGnMAHv8A//8A//8A//8A//8A//8A/wAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A
//8A//8A//8A/////wAAALAgcLAgcLAgcLAgcLAgcLAgcLAgcAAAAP////8A//8A//8A//8A//8A//8A//8A//8A//8A/////wAA
AAAACAAAdgAAsQAAzwAA9wAA/wAA/wAA/wAAwwAAiwAAsQAA5AAA/wAA/wAAaCQkDv////8A//8A//8A//8A//8A//8A//YAAP8D
Av8XFf8pJv8wLv8oJv8WFf8DA/UAAOUAAM8AALUAAJwADIUAF3UAHv8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A//8A//8A////
/wAAALAgcLAgcLAgcLAgcLAgcLAgcLAgcAAAAP////8A//8A//8A//8A//8A//8A//8A//8A//8A/////wAAAAAACAAAdAAArgAA
zAAA9wAA/wAA/wAA/wAAwwAAiwAAsQAA5AAA/wAA/wAAaCQkDv////8A//8A//8A//8A//8A//8A//4AAP8XFf84Nv9WU/9iYP9W
U/84Nf8XFf4AAO4AANkAAL8AAKQACIwAE3kAHf8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A//8A//8A//8A/////wAAALAgcLAg
cLAgcLAgcLAgcAAAAP////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A/////wAAAAAACAAAdAAArgAAzAAA9wAA/wAA/wAA
/wAAwwAAiwAAsQAA5AAA/wAA/wAAaCQkDv////8A//8A//8A//8A//8A//8A//8GBf8oJ/9WU/+Hhv+/v/+Hhv9VVP8oJv8GBfQA
AN8AAMQAAKkABo8AEXsAHP8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A//8A//8A//8A/////wAAALAgcLAgcLAgcLAgcLAgcAAA
AP////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A/////wAAAAAACAAAdAAArgAAzAAA9wAA/wAA/wAA/wAAwwAAiwAAsQAA
5AAA/wAA/wAAaCQkDv////8A//8A//8A//8A//8A//8A//8JCf8wLf9iYP/Av/7+/v/Av/9iYP8wLv8JCPYAAOEAAMcAAKsABZEA
EXwAG/8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A//8A//8A//8A//8A/////wAAALAgcLAgcLAgcAAAAP////8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A/////wAAAAAACAAAdAAArgAAzAAA9wAA/wAA/wAA/wAAyQAAmgAAxAAA7QAA/wAA/wAAaCQk
Dv////8A//8A//8A//8A//8A//8A//8A//8oJv9VVP+Hhv/Av/+Hhf9WU/8oJ/8GBvQAAN8AAMUAAKkABpAAEf8A//8A//8A//8A
//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAP8A//8A//8A//8A//8A//8A//8A//8A/////wAAALAgcLAgcLAgcAAAAP////8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A/////wAAAAAACAAAdgAAsQAAzwAA9wAA/wAA/wAA/wAAugAAdAAAlAAAvwAA3wAA6gAAVCQkEf////8A//8A//8A
//8A//8A//8A//8A//8WFf83Nf9VU/9iYP9WVP84Nv8XFf4AAO4AANkAAL8AAKQACIwAFP8A//8A//8A//8A//8A//8A//8A/wAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A
//8A//8A//8A//8A//8A//8A//8A/////wAAALAgcAAAAP////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A////
/wAAAAAACQAAfgAAvwAA2gAA+QAA/wAA/wAA/wAAiwAAAAAAAgAACAAADwAADwAAACQkI/////8A//8A//8A//8A//8A//8A//8A
//8A//8XFf8pJ/8wLf8oJv8XFf8DAvUAAOUAAM8AALYAAJwADP8A//8A//8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A//8A//8A
//8A//8A//8A/////wAAALAgcAAAAP////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A/////wAAAAAAAAAAJQAA
PgAATwAAZAAAawAAagAAcQAAKhISB/////////////////////////////8A//8A//8A//8A//8A//8A//8A//8A//8A//8GBf8K
CP8GBf0AAPYAAOkAANcAAMEAAKoABv8A//8A//8A//8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
/////wAAAP////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A/////wAAAAAAAB0dEycnGSQkEyQkDyQkDSQk
DiQkDA8PBAAAAP////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//QAAO4AAOUAANcA
AMUAAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A/////wAAAP////8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A/////wAAABUVFP//////////////////////////////////////
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//////8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//////////////////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A
/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAElM
AQEDAAMAAQAYABgA/////yEA//////////9CTTYAAAAAAAAANgAAACgAAABgAAAAMAAAAAEAGAAAAAAAADYAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAD///////////////////////////////////////////////////////////////////////////////////////////8A
AAAAAAAAAAD///////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8A
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAAAAAAAAAAAAAD///8U
FBoaGhn///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHCwIHCwIHCw
IHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHAAAAD///8AAAAAAAAAAAAAAAD///8ZGR8ZGRn///8AAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHCwIHCwIHCwIHCwIHCwIHCw
IHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHAAAAD///8AAAAAAAAAAAAAAAAAAAD///8aGiAZGRj///8AAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCw
IHCwIHCwIHCwIHCwIHAAAAD///8AAAAAAAAAAAAAAAAAAAD///8aGiAZGRj///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHAA
AAD///8AAAAAAAAAAAAAAAAAAAAAAAD///8AAAAAAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACQABKMABSGABd+ABp3AB0AAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHAAAAD///8AAAAAAAAA
AAAAAAAAAAAAAAD///8AAAAAAAD///8AAAAAAAAAAAAAAAD///////////////////////////////////////////8AAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAACqAAarAAWpAAalAAicAAySABGHABZ8ABpzAB4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAD///8AAACwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHAAAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAD/
//8AAAAAAAD///////////////////////8fHx4AAAAAAAEAAAQAAAoAAA8AAA8AAAAkJCP///8AAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAC/AADFAADHAADFAAC/AAC2AACpAAWaAA2MABR+ABpzAB4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8A
AACwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHCwIHAAAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAAAAAAdHRMn
JxkkJBMkJA8nJw8fHwoFBQAAADYAAIYAAJgAAL8AAN8AAOoAAFQkJBH///8AAAAAAAAAAAAAAAAAAAAAAAAAAADPAADZAADfAADh
AADfAADZAADPAADBAACwAAOeAAuMABR8ABtzAB4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHCwIHCw
IHCwIHCwIHCwIHCwIHCwIHCwIHAAAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAAAAAAAACUAAD4AAE8AAGQAAGkA
AGsAAHoAAIsAAKUAAMYAAO0AAP8AAP8AAGgkJA7///8AAAAAAAAAAAAAAAAAAAAAAAAAAADkAADuAAD0AAD1AAD1AADvAADlAADX
AADFAACwAAKbAAyHABZ3ABwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHCwIHCwIHCwIHCwIHCwIHCw
IHCwIHCwIHAAAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAAAAAAAAH4AAL8AANoAAPkAAP8AAP8AAP8AAMMAAIsA
ALEAAOQAAP8AAP8AAGgkJA7///8AAAAAAAAAAAAAAAAAAAAAAADoAAD2AAD9AAD/BQX/CQj/BgX+AAD1AADpAADXAADBAACpAAWS
ABF/ABpzAB4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHCwIHCwIHCwIHCwIHCwIHCwIHAAAAD///8A
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAAAAAgAAHYAALEAAM8AAPcAAP8AAP8AAP8AAMMAAIsAALEAAOQAAP8AAP8A
AGgkJA7///8AAAAAAAAAAAAAAAAAAAAAAAD2AAD/AwL/FxX/KSb/MC7/KCb/FhX/AwP1AADlAADPAAC1AACcAAyFABd1AB4AAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHCwIHCwIHCwIHCwIHCwIHCwIHAAAAD///8AAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAD///8AAAAAAAgAAHQAAK4AAMwAAPcAAP8AAP8AAP8AAMMAAIsAALEAAOQAAP8AAP8AAGgkJA7///8AAAAA
AAAAAAAAAAAAAAAAAAD+AAD/FxX/ODb/VlP/YmD/VlP/ODX/FxX+AADuAADZAAC/AACkAAiMABN5AB0AAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHCwIHCwIHCwIHCwIHAAAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAD///8AAAAAAAgAAHQAAK4AAMwAAPcAAP8AAP8AAP8AAMMAAIsAALEAAOQAAP8AAP8AAGgkJA7///8AAAAAAAAAAAAAAAAAAAAA
AAD/BgX/KCf/VlP/h4b/v7//h4b/VVT/KCb/BgX0AADfAADEAACpAAaPABF7ABwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAD///8AAACwIHCwIHCwIHCwIHCwIHAAAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAAAAAgA
AHQAAK4AAMwAAPcAAP8AAP8AAP8AAMMAAIsAALEAAOQAAP8AAP8AAGgkJA7///8AAAAAAAAAAAAAAAAAAAAAAAD/CQn/MC3/YmD/
wL/+/v7/wL//YmD/MC7/CQj2AADhAADHAACrAAWRABF8ABsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/
//8AAACwIHCwIHCwIHAAAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAAAAAgAAHQAAK4AAMwAAPcA
AP8AAP8AAP8AAMkAAJoAAMQAAO0AAP8AAP8AAGgkJA7///8AAAAAAAAAAAAAAAAAAAAAAAAAAAD/KCb/VVT/h4b/wL//h4X/VlP/
KCf/Bgb0AADfAADFAACpAAaQABEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHCwIHCw
IHAAAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAAAAAgAAHYAALEAAM8AAPcAAP8AAP8AAP8AALoA
AHQAAJQAAL8AAN8AAOoAAFQkJBH///8AAAAAAAAAAAAAAAAAAAAAAAAAAAD/FhX/NzX/VVP/YmD/VlT/ODb/FxX+AADuAADZAAC/
AACkAAiMABQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHAAAAD///8AAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAAAAAkAAH4AAL8AANoAAPkAAP8AAP8AAP8AAIsAAAAAAAIAAAgAAA8A
AA8AAAAkJCP///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/FxX/KSf/MC3/KCb/FxX/AwL1AADlAADPAAC2AACcAAwAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAACwIHAAAAD///8AAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAD///8AAAAAAAAAACUAAD4AAE8AAGQAAGsAAGoAAHEAACoSEgf///////////////////////////8A
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/BgX/Cgj/BgX9AAD2AADpAADXAADBAACqAAYAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAD///8AAAAAAAAdHRMnJxkkJBMkJA8kJA0kJA4kJAwPDwQAAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAD0AADuAADlAADXAADFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAAV
FRT///////////////////////////////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///////////////8AAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/
//8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABCTT4AAAAAAAAAPgAAACgAAABgAAAAMAAAAAEAAQAAAAAAQAIAAAAAAAAAAAAA
AAAAAAAAAAAAAAAA////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAP///////////wAAAAAAAcP//////wAAAIAAA8P//////wAAAIAAA8P//////wAAAMAAB8P///+APwAAAMAAB8P///wABwAA
AOAAD8P///wABwAAAOAAD8PAB/wAPwAAAPAAH8AAB/gA/wAAAPAAH8AAB/AAfwAAAPgAP8AAB/AAfwAAAPgAP8AAB+AAPwAAAPwA
f8AAB+AAPwAAAPwAf8AAB+AAPwAAAP4A/8AAB+AAPwAAAP4A/8AAB+AAPwAAAP8B/8AAB/AAfwAAAP8B/8AAB/AAfwAAAP+D/8AA
B/gA/wAAAP+D/8AAB/wB/wAAAP/H/8AB//8H/wAAAP/H/8AB/////wAAAP/v/8P//////wAAAP/v/////////wAAAAsAAAAPAAAA
PwIAAIpKvCA02GxBkT11TtC0wVcPUHVycGxlIFRyaWFuZ2xlBAAAABYAAAA/AgAAWkq8IDTYbEGRPXVO0LTBVwhSZWQgRmxhZwoA
AAAKAAAAPwIAAG5KvCA02GxBkT11TtC0wVcIQmx1ZSBEb3QCAAMAAAABAAAAAAD4NhJIyJOaawEAAAAAAAAAAAABAAFBEScAAA0B
BAAAAAAAAAAAADYAUHVycGxlIFRyaWFuZ2xlAAEAAAAAAJf2L2oAAAAAo/YvagAAAAAAAAAAAAAAAAAAAAAAAP//BQBBcmlhbIpK
vCA02GxBkT11TtC0wVcAAAAAAAAAAAoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAD4KRJIsIybawEAAAAAAAAAAAAC
AAFBEicAAA0BBAAAAAAAAAAAAC8AUmVkIEZsYWcAAQAAAAAApvYvagAAAACs9i9qAAAAAAAAAAAAAAAAAAAAAAAA//8FAEFyaWFs
Wkq8IDTYbEGRPXVO0LTBVwAAAAAAAAAACgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAANQ4EkiIpJxrAQAAAAAAAAAA
AAMAAUETJwAADQEEAAAAAAAAAAAALwBCbHVlIERvdAABAAAAAAC69i9qAAAAAL/2L2oAAAAAAAAAAAAAAAAAAAAAAAD//wUAQXJp
YWxuSrwgNNhsQZE9dU7QtMFXAAAAAAAAAAAKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAAABAAAAAAAAAADAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
"""

def _template_bytes():
    return base64.b64decode(_TEMPLATE_B64)






# KMZ folder -> template stream name (Stanley-based template).
LAYER_STREAMS = {
    'AGMs':       'Final AGMs68',
    'Centerline': 'Centerline41 (4)',
    'Notes':      'notes (26)',
    'Access':     'Combined Access Files1 (11)',
}

S23 = 1 << 23
S31 = 1 << 31


def enc(lon, lat):
    return ((round(lon * S23) + S31) & 0xFFFFFFFF,
            (S31 - round(lat * S23)) & 0xFFFFFFFF)


# --- colour names -> RGB (DeLorme stores COLORREF bytes R,G,B,0) ----------
def kml_color_to_rgb(kml):
    if not kml or len(kml) < 8:
        return None
    try:
        bb = int(kml[2:4], 16); gg = int(kml[4:6], 16); rr = int(kml[6:8], 16)
        return (rr, gg, bb)
    except ValueError:
        return None


# --- imagelist colour sniffing: which symbol value is which shape ---------
def _detect_symval_shapes(layer):
    bm = layer.find(b'BM')
    if bm < 0:
        return {}
    try:
        w = struct.unpack_from('<i', layer, bm + 18)[0]
        h = struct.unpack_from('<i', layer, bm + 22)[0]
        bpp = struct.unpack_from('<H', layer, bm + 28)[0]
        off = struct.unpack_from('<I', layer, bm + 10)[0]
    except struct.error:
        return {}
    if bpp != 24 or w <= 0 or h <= 0:
        return {}
    row = ((w * bpp + 31) // 32) * 4
    pix = layer[bm + off: bm + off + row * h]
    shapes = {}
    ntiles = w // 24
    for t in range(ntiles):
        purple = red = blue = 0
        for y in range(h):
            base = y * row
            for x in range(t * 24, min((t + 1) * 24, w)):
                o = base + x * 3
                if o + 2 >= len(pix):
                    continue
                b, g, r = pix[o], pix[o + 1], pix[o + 2]
                if (r, g, b) == (255, 0, 255):      # magenta = transparent
                    continue
                if b > 140 and r < 100 and g < 110:    # blue (circle)
                    blue += 1
                elif r > 100 and b > 120 and g < 110:  # purple (triangle)
                    purple += 1
                elif r > 150 and g < 100 and b < 100:   # red (flag)
                    red += 1
        best = max((blue, 'circle'), (purple, 'triangle'), (red, 'flag'))
        if best[0] > 0:
            shapes[t + 1] = best[1]
    return shapes


# ----------------------------------------------------------------------------
# Point layer template (AGMs / Notes)
# ----------------------------------------------------------------------------
class PointLayerTemplate:
    MARKER = b'\x01\x00\x00\x00\x00\x00'
    SUFFIX_LEN = {True: 97, False: 81}

    def __init__(self, data, has_symbol):
        self.has_symbol = has_symbol
        self.data = data
        coord0, count = self._bootstrap(data)
        self.count_off = coord0 - 10
        self.prefix = data[:coord0 - 6]
        self.sig = data[coord0 + 8:coord0 + 10]
        suffix_len = self._suffix_len(data, coord0)
        self.suffix_len = suffix_len

        # capture one template per symbol signature
        self.templates = {}          # sigkey -> dict
        self.default_sig = None
        pos = coord0
        for k in range(count):
            sigkey = self._obj_sig(data, pos)
            label_start = pos + 0x28
            lab_end = data.index(b'\x00', label_start)
            if sigkey not in self.templates:
                self.templates[sigkey] = dict(
                    head=bytearray(data[pos - 6:label_start]),
                    suffix=data[lab_end + 1:lab_end + 1 + suffix_len],
                    lenfield0=struct.unpack_from('<H', data, pos + 0x26)[0],
                    label0_len=lab_end - label_start,
                    oid0=struct.unpack_from('<I', data, pos + 0x16)[0],
                )
            if self.default_sig is None:
                self.default_sig = sigkey
            pos = lab_end + 1 + suffix_len + 6
        self.footer = data[pos - 6:]
        self.symval_shapes = _detect_symval_shapes(data) if has_symbol else {}

    def _obj_sig(self, data, pos):
        symval = struct.unpack_from('<H', data, pos + 0x12)[0]
        try:
            le = data.index(b'\x00', pos + 0x28)
            ap = data.find(b'Arial', le)
            builtin = struct.unpack_from('<H', data, ap - 4)[0] if ap > 0 else 0xFFFF
        except ValueError:
            builtin = 0xFFFF
        if builtin != 0xFFFF:
            return ('B', builtin)
        return ('C', symval)

    def _suffix_len(self, data, coord0):
        lab_end = data.index(b'\x00', coord0 + 0x28)
        nxt = self._next_obj(data, lab_end + 1)
        if nxt is not None:
            return nxt - 6 - (lab_end + 1)
        return self.SUFFIX_LEN[self.has_symbol]

    @staticmethod
    def _geographic(data, i):
        if i + 8 > len(data):
            return False
        X = struct.unpack_from('<I', data, i)[0]
        Y = struct.unpack_from('<I', data, i + 4)[0]
        if not (0x40000000 < X < 0x60000000 and 0x60000000 < Y < 0x70000000):
            return False
        lon = (X - S31) / S23
        lat = (S31 - Y) / S23
        return -180 < lon < 180 and -85 < lat < 85

    def _bootstrap(self, data):
        i = 0
        while True:
            j = data.find(self.MARKER, i)
            if j < 0:
                raise ValueError("no point objects found")
            coord0 = j + 6
            if self._geographic(data, coord0) and coord0 >= 10:
                count = struct.unpack_from('<I', data, coord0 - 10)[0]
                if 0 < count < 1_000_000 and self._validate(data, coord0, count):
                    return coord0, count
            i = j + 1

    def _next_obj(self, data, frm):
        i = frm
        while True:
            j = data.find(self.MARKER, i)
            if j < 0:
                return None
            c = j + 6
            if self._geographic(data, c) and data[c + 8:c + 10] == self.sig:
                return c
            i = j + 1

    def _validate(self, data, coord0, count):
        sig = data[coord0 + 8:coord0 + 10]
        self.sig = sig
        try:
            lab0_end = data.index(b'\x00', coord0 + 0x28)
        except ValueError:
            return False
        nxt = self._next_obj(data, lab0_end + 1)
        suffix_len = (nxt - 6 - (lab0_end + 1)) if nxt else None
        pos = coord0
        for k in range(count):
            if data[pos - 6:pos] != self.MARKER or not self._geographic(data, pos):
                return False
            if data[pos + 8:pos + 10] != sig:
                return False
            try:
                le = data.index(b'\x00', pos + 0x28)
            except ValueError:
                return False
            if suffix_len is None:
                return k == count - 1
            pos = le + 1 + suffix_len + 6
        return True

    def build(self, points, sig_fn):
        out = bytearray(self.prefix)
        struct.pack_into('<I', out, self.count_off, len(points))
        oid = self.templates[self.default_sig]['oid0']
        for pm in points:
            sigkey = sig_fn(pm) if self.has_symbol else self.default_sig
            tpl = self.templates.get(sigkey) or self.templates[self.default_sig]
            lon, lat = pm.coords[0]
            X, Y = enc(lon, lat)
            label = (pm.name or '').encode('latin1', 'replace')
            blob = bytearray(tpl['head'])
            struct.pack_into('<I', blob, 6 + 0x00, X)
            struct.pack_into('<I', blob, 6 + 0x04, Y)
            if self.has_symbol and isinstance(sigkey, tuple) and sigkey[0] == 'C':
                struct.pack_into('<H', blob, 6 + 0x12, sigkey[1])
            struct.pack_into('<I', blob, 6 + 0x16, oid & 0xFFFFFFFF)
            lenf = tpl['lenfield0'] + (len(label) - tpl['label0_len'])
            struct.pack_into('<H', blob, 6 + 0x26, lenf & 0xFFFF)
            out += blob + label + b'\x00' + tpl['suffix']
            oid += 1
        out += self.footer
        struct.pack_into('<I', out, 0, len(out) - 4)
        return bytes(out)


# ----------------------------------------------------------------------------
# Polyline layer template (Centerline / Access)
# ----------------------------------------------------------------------------
class LineLayerTemplate:
    LINECOUNT_OFF = 0x14

    def __init__(self, data):
        n_lines = struct.unpack_from('<I', data, self.LINECOUNT_OFF)[0]
        first = self.LINECOUNT_OFF + 4
        self.prefix = data[:first]
        pos = first
        head0 = None
        for li in range(n_lines):
            vc_off = self._vtxcount_off(data, pos)
            if head0 is None:
                head0 = data[pos:vc_off]
            n = struct.unpack_from('<I', data, vc_off)[0]
            pos = vc_off + 4 + n * 16
        self.footer = data[pos:]
        self.linehead = bytearray(head0)
        self.color_rel = self._color_rel(self.linehead)
        self.oid_rel = 4
        self.ts_rels = self._timestamp_rels(self.linehead)

    @staticmethod
    def _vtxcount_off(data, line_start):
        i, end = line_start, len(data) - 16
        while i < end:
            if data[i:i + 2] == b'\x02\x00':
                vc = struct.unpack_from('<I', data, i + 2)[0]
                if data[i + 6:i + 12] == b'\x01\x00\x00\x00\x00\x00' and 0 < vc < 5_000_000:
                    return i + 2
            i += 1
        raise ValueError("vtxcount not found")

    @staticmethod
    def _color_rel(head):
        # colour is the 4 bytes immediately before the line-style field 03 00 00 00
        idx = bytes(head).rfind(b'\x03\x00\x00\x00')
        return idx - 4 if idx >= 4 else None

    @staticmethod
    def _timestamp_rels(head):
        # The real per-line timestamps sit after the marker(4)+oid(4) header.
        # Scanning from 0 would false-match the '..00 41 <oid>' marker bytes and
        # corrupt the object type, so start past the header.
        offs = []
        for i in range(12, len(head) - 8):
            v = struct.unpack_from('<I', head, i)[0]
            if 0x60000000 < v < 0x80000000 and head[i + 4:i + 8] == b'\x00\x00\x00\x00':
                offs.append(i)
        return offs

    def build(self, lines, color_fn):
        out = bytearray(self.prefix)
        struct.pack_into('<I', out, self.LINECOUNT_OFF, len(lines))
        oid = 0x7626
        ts = int(time.time())
        for pm in lines:
            head = bytearray(self.linehead)
            if self.color_rel is not None:
                r, g, b = color_fn(pm)
                head[self.color_rel:self.color_rel + 4] = bytes((r, g, b, 0))
            struct.pack_into('<I', head, self.oid_rel, oid & 0xFFFFFFFF)
            for off in self.ts_rels:
                struct.pack_into('<I', head, off, ts & 0xFFFFFFFF)
            verts = bytearray()
            for (lon, lat) in pm.coords:
                X, Y = enc(lon, lat)
                verts += b'\x01\x00\x00\x00\x00\x00' + struct.pack('<I', X) + struct.pack('<I', Y) + b'\x00\x00'
            out += head + struct.pack('<I', len(pm.coords)) + verts
            oid += 1
        out += self.footer
        struct.pack_into('<I', out, 0, len(out) - 4)
        return bytes(out)


# ----------------------------------------------------------------------------
# AGM icon -> symbol signature
# ----------------------------------------------------------------------------
def make_agm_sig_resolver(agm_template):
    shapes = agm_template.symval_shapes        # {symval: 'triangle'/'flag'/'circle'}
    shape_to_symval = {v: k for k, v in shapes.items()}
    tri = shape_to_symval.get('triangle')
    flag = shape_to_symval.get('flag')
    circle = shape_to_symval.get('circle')

    def resolve(pm):
        icon = (pm.icon or '').lower()
        if 'triangle' in icon and tri is not None:
            return ('C', tri)
        if 'flag' in icon and flag is not None:
            return ('C', flag)
        if ('circle' in icon or 'blu' in icon or 'paddle' in icon or 'dot' in icon) and circle is not None:
            return ('C', circle)
        return agm_template.default_sig
    return resolve


# ----------------------------------------------------------------------------
# Top level
# ----------------------------------------------------------------------------
def _template_bytes():
    return base64.b64decode(_TEMPLATE_B64)


def generate_dmt(kmz_path_or_bytes):
    folders = parse_kmz(kmz_path_or_bytes)
    tpl = OLE(_template_bytes())
    streams = tpl.stream_map()
    new_streams = {}

    for folder_name, stream_name in LAYER_STREAMS.items():
        items = folders.get(folder_name, [])
        tdata = streams[stream_name]
        if folder_name in ('AGMs', 'Notes'):
            pts = [p for p in items if p.kind == 'point']
            if not pts:
                continue
            t = PointLayerTemplate(tdata, has_symbol=(folder_name == 'AGMs'))
            if folder_name == 'AGMs':
                resolver = make_agm_sig_resolver(t)
            else:
                resolver = None
            new_streams[stream_name] = t.build(pts, resolver)
        else:
            lns = [p for p in items if p.kind == 'line']
            if not lns:
                continue
            t = LineLayerTemplate(tdata)

            def color_fn(pm, fld=folder_name):
                rgb = kml_color_to_rgb(pm.color)
                if rgb:
                    return rgb
                return (255, 0, 0) if fld == 'Centerline' else (0, 0, 255)
            new_streams[stream_name] = t.build(lns, color_fn)

    # Rename layers to clean names + rewrite the Annotate index streams.
    rename = {
        'Final AGMs68': 'AGMs',
        'Combined Access Files1 (11)': 'Access',
        'notes (26)': 'Notes',
        'Centerline41 (4)': 'Centerline',
    }
    order = ['Centerline41 (4)', 'notes (26)',
             'Combined Access Files1 (11)', 'Final AGMs68']
    layer_names = [rename[o] for o in order]
    new_streams['Annotate.Filenames'] = _filenames_stream('AGMs', layer_names)
    new_streams['Annotate.ActiveFilenames'] = _active_stream('AGMs')
    return write_ole(tpl, new_streams, rename=rename)


DRAW_DIR = r'C:\DeLorme Docs\Draw'


def _lp(s):
    b = s.encode('latin1', 'replace')
    return struct.pack('<I', len(b)) + b


def _filenames_stream(active_name, layer_names):
    out = bytearray()
    out += struct.pack('<I', 1)
    out += _lp(DRAW_DIR + '\\' + active_name + '.an1')
    out += struct.pack('<I', len(layer_names))
    for nm in layer_names:
        out += _lp(nm) + struct.pack('<I', 1)
    return bytes(out)


def _active_stream(active_name):
    return struct.pack('<I', 1) + _lp(DRAW_DIR + '\\' + active_name + '.an1')


# ===========================================================================
# PART 3 : glue helpers (refactored from the original Streamlit UIs)
# ===========================================================================
def read_sheets(file_like):
    """Read the seed workbook and return (AGMs, Access, Centerline, Notes) frames."""
    df_dict = pd.read_excel(file_like, sheet_name=None)
    normalized = {k.strip().upper(): v for k, v in df_dict.items()}

    def get_sheet(*names):
        for n in names:
            if not n:
                continue
            df = normalized.get(n.strip().upper())
            if df is not None and not df.empty:
                return df
        return None

    return (get_sheet("AGMS", "AGM"),
            get_sheet("ACCESS"),
            get_sheet("CENTERLINE"),
            get_sheet("NOTES"))


def build_kmz(df_agms, df_access, df_center, df_notes):
    """Build the KMZ (bytes) exactly as the original XLSX-To-KMZ Generator does."""
    kml = simplekml.Kml()

    # Gibson logo on every output
    add_gibson_logo_overlay(kml)

    # Notes hide flags
    notes_flags_by_name = {}
    hide_col = None
    if df_notes is not None:
        for c in df_notes.columns:
            if str(c).strip().lower() == "hidenameuntilmouseover":
                hide_col = c
                break
        for _, row in df_notes.iterrows():
            nm = str(safe_str(row.get("Name")) or "").strip()
            hide_flag = True
            if hide_col:
                v = row.get(hide_col)
                if pd.notna(v) and str(v).strip().lower() in ("0", "false", "no", "n", "f"):
                    hide_flag = False
                elif pd.notna(v) and str(v).strip().lower() in ("1", "true", "yes", "y", "t"):
                    hide_flag = True
            notes_flags_by_name[nm] = hide_flag

    # AGMs
    if df_agms is not None:
        folder = kml.newfolder(name="AGMs")
        for _, row in df_agms.iterrows():
            add_agm_point(folder, row)

    # Access (keeps LineStringColor)
    if df_access is not None:
        folder = kml.newfolder(name="Access")
        created = add_lines_with_autosplit(folder, df_access, color_col="LineStringColor", split_jump_m=5000.0)
        if not created:
            for _, row in df_access.iterrows():
                add_access_point(folder, row)

    # Centerline (split on big jumps so it won't connect distant blocks)
    if df_center is not None:
        folder = kml.newfolder(name="Centerline")
        created = add_lines_with_autosplit(folder, df_center, color_col="LineStringColor", split_jump_m=5000.0)
        if not created:
            for _, row in df_center.iterrows():
                add_access_point(folder, row)

    # Notes
    if df_notes is not None:
        folder = kml.newfolder(name="Notes")
        for _, row in df_notes.iterrows():
            add_note_point(folder, row)

    # Build + inject hover styles for Notes only
    raw_kml = kml.kml().encode("utf-8")
    modified_kml = inject_hover_stylemaps_for_notes_with_flags(
        raw_kml,
        notes_flags_by_name=notes_flags_by_name,
        notes_folder_name="Notes",
    )

    # Package KMZ
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("doc.kml", modified_kml)
        zf.writestr("logo.png", base64.b64decode(GIBSON_LOGO_B64))
    return buf.getvalue()


# ===========================================================================
# PART 4 : Streamlit UI
# ===========================================================================
st.set_page_config(page_title="Seed File XLSX to KMZ and DMT",
                   page_icon=":world_map:", layout="centered")
st.title("Seed File XLSX to KMZ and DMT")
st.caption("Upload your Google Earth Seed File (.xlsx). Produces BOTH a KMZ "
           "(Earthpoint-style) and a DeLorme Street Atlas .dmt with AGMs, "
           "Access, Centerline and Notes layers — same parameters as the "
           "original two apps.")

uploaded_xlsx = st.file_uploader("Upload Google Earth Seed File (.xlsx)", type=["xlsx"])

if uploaded_xlsx is None:
    st.info("Upload a .xlsx seed file to begin.")
    st.stop()

try:
    df_agms, df_access, df_center, df_notes = read_sheets(uploaded_xlsx)
except Exception as e:
    st.error("Failed to read Excel file: %s" % e)
    st.stop()

tab1, tab2, tab3, tab4 = st.tabs(["AGMs", "Access", "Centerline", "Notes"])
with tab1:
    st.subheader("AGMs")
    st.dataframe(df_agms if df_agms is not None else pd.DataFrame())
with tab2:
    st.subheader("Access")
    st.dataframe(df_access if df_access is not None else pd.DataFrame())
with tab3:
    st.subheader("Centerline")
    st.dataframe(df_center if df_center is not None else pd.DataFrame())
with tab4:
    st.subheader("Notes")
    st.dataframe(df_notes if df_notes is not None else pd.DataFrame())

base = os.path.splitext(uploaded_xlsx.name)[0]

if st.button("Generate KMZ + DMT", type="primary"):
    with st.spinner("Building KMZ ..."):
        try:
            kmz_data = build_kmz(df_agms, df_access, df_center, df_notes)
        except Exception as e:
            st.exception(e)
            st.stop()

    with st.spinner("Building DeLorme .dmt ..."):
        try:
            dmt_data = generate_dmt(kmz_data)
        except Exception as e:
            st.exception(e)
            st.stop()

    st.success("Done! Both files are ready.")
    c1, c2 = st.columns(2)
    c1.download_button("Download KMZ", data=kmz_data,
                       file_name="%s.kmz" % base,
                       mime="application/vnd.google-earth.kmz")
    c2.download_button("Download DMT", data=dmt_data,
                       file_name="%s.dmt" % base,
                       mime="application/octet-stream")
