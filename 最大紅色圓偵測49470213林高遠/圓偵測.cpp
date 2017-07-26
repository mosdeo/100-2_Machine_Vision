#include "stdafx.h"
//100-2 機電系 機械視覺
//期中考-最大紅色圓偵測
//工教系494702123 林高遠

char FileName[20]="Circleshape.bmp";
double darrList[20][3];

int main()
{
	std::cout << "100-2 機電系 機械視覺\n";
	std::cout << "期中考 最大紅色圓偵測\n";
	std::cout << "工教系 494702123 林高遠\n";
	std::cout << "讀入 Circleshape.bmp ...\n\n";

	if(FileName[0]==NULL){
		std::cout << "找不到 Circleshape.bmp ...!\n\n";
		system("pause");
		return -1;
	}

	cv::Mat Img=cv::imread(FileName,1);
	cv::Mat ImgResult=Img.clone();
	//複製也可以用
	//Img.copyTo(ImgResult);
	//不可以用
	//ImgResult=Img;
	//會連帶處理


	/*過濾紅色(BGR排列)*/
	for(int i=0;i<Img.rows;i++){
		for(int j=0;j<Img.cols;j++){
			if(Img.at<cv::Vec3b>(i,j)[2] > (Img.at<cv::Vec3b>(i,j)[0] + Img.at<cv::Vec3b>(i,j)[1]))
			{
				ImgResult.at<cv::Vec3b>(i,j)[0]=255;
				ImgResult.at<cv::Vec3b>(i,j)[1]=255;
				ImgResult.at<cv::Vec3b>(i,j)[2]=255;
			}
			else
			{
				ImgResult.at<cv::Vec3b>(i,j)[0]=0;
				ImgResult.at<cv::Vec3b>(i,j)[1]=0;
				ImgResult.at<cv::Vec3b>(i,j)[2]=0;
			}
		}
	}

	/*轉單通道,否則不能做形態運算*/
	cv::cvtColor(ImgResult,ImgResult,cv::COLOR_BGR2GRAY,0);

	/*對灰階圖閉運算,消除會造成干擾的細線*/
	cv::erode(ImgResult,ImgResult,cv::Mat(),cv::Point(-1,-1),3);
	cv::dilate(ImgResult,ImgResult,cv::Mat(),cv::Point(-1,-1),3);

	std::vector<cv::Vec3f> circles;
	cv::HoughCircles(ImgResult,circles,CV_HOUGH_GRADIENT,
		Img.rows/64, //累加解析
		20,			//兩圓間最小距離
		200,		//Canny 門檻
		50,		//投票門檻
		5,120);		//最小最大半徑

	/*各組圓的大小排序*/
	for(int i=0;i<circles.size()-1;i++){
		for(int j=0;i+j<circles.size()-1;j++){
			if(circles[i][2]<circles[i+j][2]){
				std::swap(circles[i][0],circles[i+j][0]);
				std::swap(circles[i][1],circles[i+j][1]);
				std::swap(circles[i][2],circles[i+j][2]);
			}
		}
	}


	/*畫出欲偵測的圓*/
	for(int i=0;i<1;i++ )
	{
		cv::Point center(cvRound(circles[i][0]), cvRound(circles[i][1]));
		int radius = cvRound(circles[i][2]);
		// circle center
		circle( Img, center, 3, cv::Scalar(0,255,0), -1, 8, 0 );
		// circle outline
		circle( Img, center, radius, cv::Scalar(0,255,0), 3, 8, 0 );
	}

	std::cout<<"最大紅色圓 x="<< cvRound(circles[0][0]) <<" y=" << cvRound(circles[0][1]) << " r=" << cvRound(circles[0][2]) ;
	cv::imshow("Img",Img);
	cv::imshow("ImgResult",ImgResult);
	cv::waitKey(0);
	return 0;
}