import sys, numpy as np, cv2
sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.abspath(__file__)))
from common import render
def lines(book,pdf,y0,y1,axis='v'):
    img=render(book,pdf,100); g=cv2.cvtColor(np.array(img),cv2.COLOR_RGB2GRAY)
    H,W=g.shape; bw=(g<150).astype(np.uint8)
    if axis=='v':
        sub=bw[int(y0*H):int(y1*H)]; prof=sub.sum(0)/sub.shape[0]; n=W
    else:
        x0,x1=y0,y1; sub=bw[:,int(x0*W):int(x1*W)]; prof=sub.sum(1)/sub.shape[1]; n=H
    xs=[i for i in range(n) if prof[i]>float(__import__("os").environ.get("T","0.5"))]
    grp=[];
    for x in xs:
        if grp and x-grp[-1][-1]<=2: grp[-1].append(x)
        else: grp.append([x])
    return [round(sum(g)/len(g)/n,3) for g in grp]
b,p=int(sys.argv[1]),int(sys.argv[2]); a=sys.argv[5] if len(sys.argv)>5 else 'v'
print(lines(b,p,float(sys.argv[3]),float(sys.argv[4]),a))
