/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from 'react';
import axios from 'axios';
import Slider from 'react-slick';

import { IBannerProps } from './IBannerProps';

import 'slick-carousel/slick/slick.css';
import 'slick-carousel/slick/slick-theme.css';
import styles from './Banner.module.scss';

interface PropsBanner {
  Id: number;
  Title: string;
  Ativo: boolean;
  Banner: string;
}

const Banner:React.FC<IBannerProps> = (props: IBannerProps) => {
  const [data, setData] = React.useState<PropsBanner[]>([])
  
  React.useEffect(() =>  {
    async function getBanner() {
      const resp = await axios(`${props.site.absoluteUrl}/_api/web/lists/getbytitle('${props.description}')/items?$select=Id,Title,Banner,Ativo`).then(res => { return res.data.value})
      return setData(resp)
    }
    getBanner()
  }, [])
  
  if(!props.checkbox){
    let activeElement: PropsBanner | null = null;
    for (const elm of data) {
      if (elm.Ativo) {
        activeElement = elm;
        break;
      }
    }   
    return (
      activeElement && (
          <>
          {(() => {
            const img = JSON.parse(activeElement.Banner);
            return (
              <div className={styles.banner}>
                <img src={`${img.serverUrl}/${img.serverRelativeUrl}`} alt={activeElement.Title} />
                <h1>{activeElement.Title}</h1>
              </div>
            );
          })()}
          </>
        )
    )
  } else {
    return (
      <>
        <Slider dots slidesToShow={1} autoplay waitForAnimate={false} speed={props.secondsSlick}>
          {data.map((elm: PropsBanner, index: number) => {
            const img = JSON.parse(elm.Banner);
            return (
              <div key={index} className={styles.banner} style={{"background-color": props.color || '#00000040'} as React.CSSProperties}>
                <img src={`${img.serverUrl}/${img.serverRelativeUrl}`} alt={elm.Title} />
                <h1>{elm.Title}</h1>
              </div>
            );
          })}
        </Slider>
      </>
    );
  }
}

export default Banner;