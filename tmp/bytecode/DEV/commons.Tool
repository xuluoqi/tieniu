6f5969f816f1c2493deb8be58fb1172c ����   3 �  commons/Tool  java/lang/Object <init> ()V Code
  	   LineNumberTable LocalVariableTable this Lcommons/Tool; getPage '(Ljava/lang/Integer;)Ljava/lang/String;  java/lang/StringBuffer
  	  &<div id='page' class='paginationPage'>
     append ,(Ljava/lang/String;)Ljava/lang/StringBuffer;  <div class='pagination-item'>
    java/lang/Integer    intValue ()I " Q<a target='_self' href='#' id='page_1' onclick='toPage(1)'  class='current'>1</a> $ java/lang/StringBuilder & $<a target='_self' href='#' id='page_
 # (  ) (Ljava/lang/String;)V
 # +  , (I)Ljava/lang/StringBuilder; . ' onclick='toPage(
 # 0  1 -(Ljava/lang/String;)Ljava/lang/StringBuilder; 3 )'  class='current'> 5 </a>
 # 7 8 9 toString ()Ljava/lang/String; ; )'> = F<a target='_self' id='page_1' onclick='toPage(1)' href='#'>首页</a>  ? <a target='_self' id='page_ A )' href='#'> C </a> 
 # E  F -(Ljava/lang/Object;)Ljava/lang/StringBuilder; H "<a target='_self' onclick='toPage( J )'  href='#'>...</a>  L </div> N +<a target='_self' href='#' onclick='toPage( P ;)'  class='pagination-next' title='下一页'>下一页</a>
  7 page Ljava/lang/Integer; sb Ljava/lang/StringBuffer; i I StackMapTable getKey [   ] java/util/Random
 \ 	
 \ ` a b nextInt (I)I d char f num
 h j i java/lang/String k l equalsIgnoreCase (Ljava/lang/String;)Z
 h n o p valueOf &(Ljava/lang/Object;)Ljava/lang/String;
 # r  s (C)Ljava/lang/StringBuilder;
 h u o v (I)Ljava/lang/String;
 h x y 9 toUpperCase { 0 } A
 h  � � 
replaceAll 8(Ljava/lang/String;Ljava/lang/String;)Ljava/lang/String; � O � X � o length val Ljava/lang/String; random Ljava/util/Random; 	charOrNum choice vall Num2S (JI)Ljava/lang/String;
 h � o � (J)Ljava/lang/String; � \.
 h � �   m J len s c getBootstartPage � <div class='pagination'><ul> � <li><a href='#'onclick='toPage( � )')>上一页</a></li> � <li><a href='#' id='page_ � &' class='btn-success' onclick='toPage( � 	</a></li> � )'>下一页</a></li> � </ul></div> 
SourceFile 	Tool.java !               /     *� �    
                    	      �    � Y� L+� W+� W*� *� � �*� +!� W��=� i*� � 2+� #Y%� '� *-� /� *2� /� *4� /� 6� W� /+� #Y%� '� *-� /� *:� /� *4� /� 6� W�
����+<� W*� d=� 2+� #Y>� '� *-� /� *@� /� *B� /� 6� W�*� ���+� #Y%� '*� D-� /*� D2� /*� D4� /� 6� W*� `=� 2+� #Y>� '� *-� /� *@� /� *B� /� 6� W�*� `���+� #YG� '*� `� *I� /� 6� W*� `=� 2+� #Y>� '� *-� /� *@� /� *B� /� 6� W�*� 
`���+K� W+� #YM� '*� `� *O� /� 6� W+K� W+� Q�    
   n           "  &  0  5  =  l  �  �  �  �   �  � " #" $N #[ &z '� (� '� +� ,� -� .    >   � R S   � T U  2 o V W  � : V W  < V W � = V W  X     � " � 6+� � .=.3.� 
 	 Y     �     �ZL� \Y� ^M>� z,� _p� c� e:c� g� 5,� _p� A� a6� #Y+� m� ',� _`�� q� 6L� (e� g� � #Y+� m� ',
� _� t� /� 6L�*� ���+� wN-z|� ~N-z|� ~N-��� ~N-��� ~N-�    
   B    8  9  :  < # > - @ @ A _ C i E � : � H � I � J � K � L � M    H    � � S    � � �   � � �   � V W  # a � �  @  � W  � & � �  X    �  h \A h�  hA � $ 	 � �     �     A� #Y� �� '� 6�Z� ~N-� �66� � #Yz� '-� /� 6N�d���-�    
       X  Y  Z " [ 3 Z ? ]    4    A � �     A � W   + � �   % � W     V W  X    � " h 	 �     �    �� Y� L+�� W+� #Y�� '*� *� � � 	*� d� *�� /� 6� W*� *� � w=� i*� � 2+� #Y�� '� *�� /� *:� /� *�� /� 6� W� /+� #Y�� '� *-� /� *:� /� *�� /� 6� W�
���� �*� d=� 2+� #Y�� '� *-� /� *:� /� *�� /� 6� W�*� ���+� #Y�� '*� D�� /*� D:� /*� D�� /� 6� W*� `=� 2+� #Y�� '� *-� /� *:� /� *�� /� 6� W�*� `���+� #Y�� '*� `� *�� /� 6� W+�� W+� Q�    
   N    h  i  j > k J l O m W n � p � l � s � t � s � v+ x5 ya xn |� }� ~    4   � R S   � T U  L o V W  � : V W 2 < V W  X   Q � %      #�       #�       #� 6+� � 	.=.� 	  �    �