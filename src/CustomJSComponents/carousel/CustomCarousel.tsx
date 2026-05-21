import * as React from 'react';
import '../carousel/Carouselcss.scss';

const CustomCarousel: React.FC = () => {
  const [activeIndex, setActiveIndex] = React.useState<number>(0);
  const carouselRef = React.useRef<HTMLDivElement>(null);
  const totalSlides = 3;

  const goToNextSlide = () => {
    setActiveIndex((prevIndex) => (prevIndex + 1) % totalSlides);
  };

  const goToPrevSlide = () => {
    setActiveIndex((prevIndex) => (prevIndex - 1 + totalSlides) % totalSlides);
  };

  const handleSwipe = (event: any, direction: string) => {
    if (direction === 'left') goToNextSlide();
    if (direction === 'right') goToPrevSlide();
  };

  React.useEffect(() => {
    const carouselElement = carouselRef.current;
    if (!carouselElement) return;

    const handleTouchStart = (e: TouchEvent) => {
      const startX = e.touches[0].clientX;
      const handleTouchMove = (moveEvent: TouchEvent) => {
        const moveX = moveEvent.touches[0].clientX;
        const deltaX = startX - moveX;

        if (deltaX > 50) handleSwipe(e, 'left');
        if (deltaX < -50) handleSwipe(e, 'right');

        carouselElement.removeEventListener('touchmove', handleTouchMove);
      };

      carouselElement.addEventListener('touchmove', handleTouchMove);
    };

    carouselElement.addEventListener('touchstart', handleTouchStart);

    return () => {
      if (carouselElement) {
        carouselElement.removeEventListener('touchstart', handleTouchStart);
      }
    };
  }, []);

  return (
    <div className="container">
      <h3 className="text-center text-uppercase">Custom Bootstrap Carousel</h3>
      <div className="row">
        <div className="col-lg-10 col-md-8 col-sm-12 mx-auto my-5">
          <div
            id="carouselExampleIndicators"
            className="carousel slide"
            ref={carouselRef}
          >
            <ol className="carousel-indicators">
              {[...Array(totalSlides)].map((_, index) => (
                <li
                  key={index}
                  data-slide-to={index}
                  className={index === activeIndex ? 'active' : ''}
                  onClick={() => setActiveIndex(index)}
                ></li>
              ))}
            </ol>
            <div className="carousel-inner">
              <div className={`carousel-item ${activeIndex === 0 ? 'active' : ''}`}>
                <img
                  src="https://images.unsplash.com/photo-1561877202-53d0e24be55d?ixlib=rb-1.2.1&ixid=eyJhcHBfaWQiOjEyMDd9&auto=format&fit=crop&w=1050&q=80"
                  className="d-block w-100"
                  alt="Slide 1"
                />
              </div>
              <div className={`carousel-item ${activeIndex === 1 ? 'active' : ''}`}>
                <img
                  src="https://images.unsplash.com/photo-1561622245-4d9cd72441a8?ixlib=rb-1.2.1&auto=format&fit=crop&w=1050&q=80"
                  className="d-block w-100"
                  alt="Slide 2"
                />
              </div>
              <div className={`carousel-item ${activeIndex === 2 ? 'active' : ''}`}>
                <img
                  src="https://images.unsplash.com/photo-1508724735996-b41f69dfe2a9?ixlib=rb-1.2.1&auto=format&fit=crop&w=1156&q=80"
                  className="d-block w-100"
                  alt="Slide 3"
                />
              </div>
            </div>
            <a
              className="carousel-control-prev"
              href="#carouselExampleIndicators"
              role="button"
              onClick={goToPrevSlide}
            >
              <span className="carousel-control-prev-icon" aria-hidden="true"></span>
              <span className="sr-only">Previous</span>
            </a>
            <a
              className="carousel-control-next"
              href="#carouselExampleIndicators"
              role="button"
              onClick={goToNextSlide}
            >
              <span className="carousel-control-next-icon" aria-hidden="true"></span>
              <span className="sr-only">Next</span>
            </a>
          </div>
        </div>
      </div>
    </div>
  );
};

export default CustomCarousel;
