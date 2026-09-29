import numpy as np
from PIL import Image, ImageDraw, ImageFilter
W,H=2400,1350
def lerp(a,b,t): return a+(b-a)*t
def hex2(c): return np.array([int(c[i:i+2],16) for i in (0,2,4)],float)

def gradient(c1,c2,c3):
    y,x=np.mgrid[0:H,0:W]
    t=(x/W*0.65+y/H*0.35)
    a,b,c=hex2(c1),hex2(c2),hex2(c3)
    t=t[...,None]
    img=np.where(t<0.55, a+(b-a)*(t/0.55), b+(c-b)*((t-0.55)/0.45))
    # soft glow top-right
    g=np.exp(-(((x-W*0.82)/(W*0.35))**2+((y-H*0.15)/(H*0.55))**2))[...,None]
    img=img+ (np.array([90,140,255])-img)*g*0.22
    return Image.fromarray(np.clip(img,0,255).astype('uint8'))

def waves(img,color,alpha,n=14,y0=0.62,amp=0.10,spread=0.012,x_from=0.0,width=3):
    ov=Image.new('RGBA',img.size,(0,0,0,0)); d=ImageDraw.Draw(ov)
    xs=np.linspace(x_from*W,W,400)
    for i in range(n):
        ph=i*0.18
        ys=H*(y0+spread*i) + H*amp*np.sin(xs/W*2*np.pi*0.9+ph) + H*0.05*np.sin(xs/W*2*np.pi*2.1+ph*1.7)
        pts=list(zip(xs,ys))
        a=int(alpha*(0.35+0.65*(1-abs(i-n/2)/(n/2))))
        d.line(pts,fill=color+(a,),width=width)
    ov=ov.filter(ImageFilter.GaussianBlur(0.8))
    return Image.alpha_composite(img.convert('RGBA'),ov).convert('RGB')

# Title / dark slides
t=gradient('071537','0F2F7A','2A5FD6')
t=waves(t,(255,255,255),60,n=18,y0=0.55,amp=0.09,spread=0.011,x_from=0.30)
t.save('build/img/bg_dark.png')
# Closing / exec variant: calmer
t2=gradient('081739','11337F','234FB8')
t2=waves(t2,(255,255,255),32,n=12,y0=0.70,amp=0.07,spread=0.012,x_from=0.45)
t2.save('build/img/bg_dark2.png')

# Content slides: white with faint blue wash at top right + faint waves bottom right
y,x=np.mgrid[0:H,0:W]
base=np.ones((H,W,3))*255
g=np.exp(-(((x-W*1.0)/(W*0.45))**2+((y-H*0.0)/(H*0.55))**2))[...,None]
tint=np.array([226,235,251],float)
img=base+(tint-base)*g*0.9
c=Image.fromarray(np.clip(img,0,255).astype('uint8'))
c=waves(c,(47,95,214),22,n=10,y0=0.02,amp=0.05,spread=0.010,x_from=0.62,width=2)
c.save('build/img/bg_content.png')

# Crops
ga=Image.open('src_x/ppt/media/image1.png'); ga.crop((34,24,1286,828)).save('build/img/ga_panel.png')
g2=Image.open('src_x/ppt/media/image2.png')
tiles={'howto':(18,16,506,778),'mw_account':(509,16,997,778),'ali':(1000,16,1488,778),'mw_mention':(1491,16,1979,778),'musheera':(1491,781,1979,1533)}
for k,b in tiles.items(): g2.crop(b).save(f'build/img/tile_{k}.png')
print('ok')
